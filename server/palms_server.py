#!/usr/bin/env python3
"""Palms Server: the server every Palms LAN Center at the party connects to, for chat, the schedule,
votes, tournaments, the Game Library list and the file page."""

import asyncio
import collections
import configparser
import datetime
import hmac
import json
import os
import pathlib
import random
import re
import secrets
import socket
import struct
import subprocess
import sys

VERSION = '0.5.2'
HERE = pathlib.Path(__file__).resolve().parent
CONFIG_FILE = HERE / 'palms.ini'
DATA_DIR = HERE / 'data'
PCS_FILE = DATA_DIR / 'pcs.json'
SCHED_FILE = DATA_DIR / 'schedule.json'
TPL_FILE = DATA_DIR / 'templates.json'
MAPS_FILE = DATA_DIR / 'seatmaps.json'
FOOD_FILE = DATA_DIR / 'food.json'
CUPS_FILE = DATA_DIR / 'cups.json'
FILES_DIR = HERE / 'files'

NICK_RE = re.compile(r'^[A-Za-z0-9_\-]{2,12}$')
SEAT_RE = re.compile(r'^[A-Z]{1,2}[0-9]{2}$')
PREFIX_RE = re.compile(r'^[A-Z]{1,2}$')
CHIMES = ('dingdong', 'triple', 'horn', 'none')
TIME_RE = re.compile(r'^([0-9]{1,2}):([0-9]{2})$')
EVENT_ICONS = ('trophy', 'crosshair', 'rocket', 'flag', 'castle', 'planet', 'tank', 'sword', 'bomb',
               'pizza', 'burger', 'moon', 'star', 'door', 'megaphone', 'joystick', 'calendar', 'disc')
MAX_VOTE_OPTIONS = 6
MAX_SUGGESTIONS = 12       # games suggested for the next vote, all users together
MAX_BACKED = 3             # suggestions one user can back at a time
MAX_NOTE = 300             # event details
MAX_CUPS = 8               # tournaments at once
MAX_CUP_PLAYERS = 64
CUP_STATES = ('signup', 'drawn', 'running', 'done')
MAX_LINE = 4096           # long enough for a whole schedule template on one line
MAX_TEXT = 400
HISTORY_LINES = 20
REAL_CLIENT_TIMEOUT = 180   # seconds of silence before a real client is dropped
PARTY_HOURS = 16            # "here since" and "gone home" are forgotten after this long
MAX_GONE = 30               # how many people who've left Who's Here lists
PING_SECONDS = 10           # how often the server checks each PC's ping
PAK_SECONDS = 10            # how often the server asks the pak server for its game list

HELP_TEXT = [
    'Commands:',
    'LOGIN|nick|seat          sign in, for example LOGIN|Player1|A01',
    'CHAT|message             say something to everyone',
    'WHO                      list users signed in, and users who signed out today',
    'WHOIS|nick               details for one user: seat, PC, system information',
    'SEAT|A01  or  SEAT|-     change seat, or clear it',
    'SPECS|cpu|memory|graphics|windows|screen|sound|network card|link Mb   send system information',
    'MAP                      show the active seat map',
    'PAKS                     list the games on the Cartridge Server',
    'MAPLIST  MAPGET|id  MAPSAVE|id|name|stage|columns|tables  MAPDEL|id  MAPUSE|id   host only, seat maps',
    '                         tables look like  Table A,A,4,2~Table B,B,4,2  (name, seat letter, seats per side, sides)',
    'AWAY, BUSY or BACK       set status to away, busy or online',
    'HOST|password            enable host mode (ADMIN|password also works)',
    'PA|chime|message         host only. Chime is dingdong, triple, horn or none',
    'SCHED                    show today\'s schedule',
    'SCHEDADD|19:30|icon|name host only. Icons: ' + ', '.join(EVENT_ICONS),
    'SCHEDDEL|id              host only. Remove an event (the ID is the first field in SCHED)',
    'SCHEDNOTE|id|details     host only. Event details, ~ starts a new line. Blank removes them',
    'SCHEDIN|id  SCHEDOUT|id  count yourself in or out of an event',
    'VOTEGET                  show the current vote',
    'VOTE|2                   vote for choice 2',
    'VOTESTART|5|question|choice 1|choice 2|...   host only. Minutes, then 2 to 6 choices',
    'VOTESTOP                 host only. Close the vote now',
    'SUGGEST|game             suggest a game for the next vote (or back one already suggested)',
    'UNSUGGEST|game           withdraw your suggestion',
    'SUGGDEL|game             host only. Remove a suggestion',
    'FOOD|order               place or change your food order. FOOD| on its own cancels it',
    'FOODOPEN|name  FOODCLOSE  FOODCLEAR   host only, food orders',
    'TPLLIST  TPLGET|id  TPLSAVE|id|name|line~line  TPLDEL|id  TPLAPPLY|replace|line~line   host only, schedule templates',
    'CUPS                     list the tournaments (and hear about changes)',
    'CUPJOIN|id  CUPLEAVE|id  sign up for a tournament, or withdraw',
    'CUPREPORT|id|round|match|a or b|score   report the winner of your match (round and match count from 0)',
    'CUPNEW|max|time|game|name|notes  CUPEDIT|id|max|time|game|name|notes   host only. max 0 = no limit',
    'CUPDRAW|id  CUPSTART|id  CUPOPEN|id  CUPDEL|id   host only. Draw (or redraw), start, reopen sign-ups, delete',
    'CUPADD|id|nick  CUPREMOVE|id|nick  CUPSWAP|id|nick|nick   host only, the players',
    'CUPRESULT|id|round|match|a or b|score   host only. Confirm or enter a result. a or b left blank clears it',
    'LOGOUT|break  or  LOGOUT|home   sign out (break keeps the seat)',
    'PING                     check the connection',
    'QUIT                     disconnect',
    'QUIT|update              disconnect to install an update (the LAN Center does this itself)',
]


def now_hm():
    return datetime.datetime.now().strftime('%H:%M')


def log(msg):
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg, flush=True)


def party_minutes(hm):
    """Minutes since midnight, but times before 6 AM count as after midnight of the same party."""
    h, m = (int(x) for x in hm.split(':'))
    mins = h * 60 + m
    return mins + 1440 if mins < 360 else mins


def ampm(hm):
    h, m = (int(x) for x in hm.split(':'))
    return '%d:%02d %s' % (h % 12 or 12, m, 'AM' if h < 12 else 'PM')


# Schedule templates. Each line is a time, then what's on: "12:15 PM  Game 1".
DEFAULT_TEMPLATES = [
    ('Full day', ['10:00 AM\tDoors open - Setup begins', '12:00 PM\tIntroduction & rules', '12:15 PM\tGame 1',
                  '1:00 PM\tLunch break', '2:00 PM\tGame 2', '4:00 PM\tFree play session', '6:30 PM\tDinner break',
                  '7:30 PM\tGame 3', '9:00 PM\tFree play until wrap-up', '10:00 PM\tFinal matches wrap-up',
                  '10:30 PM\tOut by 10:30 PM']),
    ('Afternoon and evening', ['1:00 PM\tDoors open - Setup begins', '2:00 PM\tIntroduction & rules', '2:15 PM\tGame 1',
                               '4:00 PM\tFree play session', '6:00 PM\tDinner break', '7:00 PM\tGame 2',
                               '8:30 PM\tTournament final', '9:30 PM\tFree play until wrap-up', '11:30 PM\tFinal matches wrap-up',
                               '12:00 AM\tOut by midnight']),
    ('Evening session', ['6:00 PM\tDoors open - Setup begins', '6:30 PM\tIntroduction & rules', '6:45 PM\tGame 1',
                         '8:00 PM\tPizza break', '8:30 PM\tGame 2', '10:00 PM\tFree play until wrap-up',
                         '11:00 PM\tOut by 11:00 PM']),
]
TPL_LINE_RE = re.compile(r'^\s*(\d{1,2})(?:[:.](\d{2}))?\s*([ap])?\.?\s*(?:m\.?)?(?![a-z])\s*[-:\t ]*\s*(.*)$', re.I)
AUTO_ICONS = [
    ('door', ('door', 'out by', 'pack up', 'setup', 'set up', 'close')),
    ('megaphone', ('intro', 'rules', 'welcome', 'announce', 'briefing', 'raffle')),
    ('burger', ('dinner', 'bbq', 'barbecue', 'burger', 'sausage')),
    ('pizza', ('lunch', 'pizza', 'snack', 'food', 'breakfast')),
    ('trophy', ('final', 'tournament', 'cup', 'prize', 'award', 'match')),
    ('joystick', ('game', 'play', 'lan')),
    ('moon', ('wrap', 'late', 'night', 'sleep', 'midnight')),
]


def parse_template_line(line):
    """'12:15 PM  Game 1' -> ('12:15', 'joystick', 'Game 1'), or None when there's no time at the start."""
    m = TPL_LINE_RE.match(line)
    if not m or (m.group(2) is None and m.group(3) is None):
        return None                       # "1v1 tournament" is not a time
    h, mins = int(m.group(1)), int(m.group(2) or 0)
    if m.group(3):
        if not 1 <= h <= 12:
            return None
        h = h % 12 + (12 if m.group(3).lower() == 'p' else 0)
    if h > 23 or mins > 59:
        return None
    text = m.group(4).strip()
    icon = None
    b = re.match(r'\[(\w+)\]\s*(.*)', text)
    if b and b.group(1).lower() in EVENT_ICONS:
        icon, text = b.group(1).lower(), b.group(2).strip()
    if not text:
        return None
    if not icon:
        low = text.lower()
        icon = next((ic for ic, words in AUTO_ICONS if any(w in low for w in words)), 'star')
    return ('%02d:%02d' % (h, mins), icon, text[:40])


# The seat map a new server starts with: three tables, four seats down each side.
DEFAULT_MAP = {'id': 1, 'name': 'Palms default', 'stage': 'Big screen and stage', 'columns': 1,
               'tables': [{'name': 'Table %s' % x, 'prefix': x, 'per_side': 4, 'sides': 2} for x in 'ABC']}


def clean(text, limit=MAX_TEXT):
    """Remove control characters and trim, so one client can't mess up another's screen."""
    text = ''.join(ch for ch in text if ch >= ' ' and ch != '\x7f')
    return text.strip()[:limit]


def broadcast_addresses():
    """The broadcast address of each IPv4 network this server is on, like 10.0.200.255 (from ip addr)."""
    try:
        out = subprocess.run(['ip', '-o', '-4', 'addr', 'show'], capture_output=True, text=True, timeout=5).stdout
    except (OSError, subprocess.SubprocessError):
        return []
    found = []
    for line in out.splitlines():
        f = line.split()
        if len(f) > 1 and f[1] == 'lo':
            continue
        if 'brd' in f:
            b = f[f.index('brd') + 1]
            if b not in found:
                found.append(b)
    return found


def fetch_text(url, limit=2000000):
    """A small text file from a web server (the pak server's catalog.txt). Runs in a thread."""
    import urllib.request
    with urllib.request.urlopen(url, timeout=5) as r:
        return r.read(limit).decode('cp1252', 'replace')


# ---------------------------------------------------------------- settings

def load_config():
    cp = configparser.ConfigParser()
    if CONFIG_FILE.exists():
        cp.read(CONFIG_FILE, encoding='utf-8')
    if 'server' not in cp:
        cp['server'] = {}
    s = cp['server']
    defaults = {'name': 'PALMS-SRV', 'tcp_port': '4098', 'beacon_port': '4099',
                'beacon_seconds': '3', 'admin_password': '', 'http_port': '8080',
                'paks_url': 'http://10.0.200.73/'}
    changed = False
    for key, value in defaults.items():
        if key not in s:
            s[key] = value
            changed = True
    if not s['admin_password']:
        s['admin_password'] = secrets.token_hex(4)
        changed = True
        log('Generated host password: %s  (saved in %s, as admin_password)' % (s['admin_password'], CONFIG_FILE))
    if changed:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            f.write('# Palms Server settings. Restart the server after making changes.\n')
            cp.write(f)
    return s


# ---------------------------------------------------------------- the file page (port 8080)
#
# A small web page every PC at the party can open, even Internet Explorer 6:
#   http://<server>:8080/                  list of files, upload form
#   /client/...                            the newest LAN Center program files
#   /client/manifest.txt                   the list palms-update.vbs downloads
#   /palms-update.vbs                      installs or updates the LAN Center on a PC
#   /palms-clean.vbs                       clean slate: wipes the games and settings off a test PC
# Downloading and uploading are open to everyone on the network. A new version of
# the LAN Center is published by uploading PalmsLAN.exe with the form on the page: the LAN
# Center's own files (CLIENT_FILES) go to the LAN Center Files, everything else to Other Files.

import html as _html
import http.server
import mimetypes
import threading
import time
import urllib.parse

MAX_UPLOAD = 200 * 1024 * 1024
SAFE_NAME = re.compile(r'[^A-Za-z0-9._()\- ]')
DOWNLOAD_ONLY = ('.exe', '.ocx', '.dll', '.vbs', '.res', '.frx', '.zip', '.bat', '.reg', '.hta')
CLIENT_LAST = ('PalmsLAN.exe',)      # fetched last, so a half-finished update never runs
# Files older versions used. Removed from the LAN Center files when the file page starts.
RETIRED = ('CartridgeMaker.hta', 'pakmaker.gif', 'palm.ico', 'wget.exe')
# Uploaded with these names, a file goes to the LAN Center Files, which every PC installs
CLIENT_FILES = ('PalmsLAN.exe', 'MSWINSCK.OCX', 'palmsget.exe', '7za.exe', 'ding.wav', 'dingdong.wav', 'horn.wav', 'triple.wav')


def exe_version(path):
    """The version in a Windows program's version resource, like 0.7.6, or '' if there isn't one.
    Visual Basic 6.0 writes Major.Minor.0.Revision there (Project > Properties > Make)."""
    try:
        data = path.read_bytes()
    except OSError:
        return ''
    # the VS_VERSION_INFO resource, then its VS_FIXEDFILEINFO, which starts with a signature
    k = data.find('VS_VERSION_INFO'.encode('utf-16-le'))
    i = data.find(b'\xbd\x04\xef\xfe', k, k + 64) if k >= 0 else -1
    if i < 0 or len(data) < i + 16:
        return ''
    ms, ls = struct.unpack_from('<II', data, i + 8)
    major, minor, build, rev = ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF
    if build:
        return '%d.%d.%d.%d' % (major, minor, build, rev)
    return '%d.%d.%d' % (major, minor, rev)


def ver_tuple(v):
    """'0.7.6' -> (0, 7, 6), for comparing versions."""
    out = []
    for part in (v or '').split('.')[:4]:
        m = re.match(r'\d+', part)
        out.append(int(m.group()) if m else 0)
    return tuple(out + [0] * (4 - len(out)))


def safe_name(raw):
    """Turn whatever the browser sent (IE6 sends the whole C:\\path) into a plain file name."""
    name = raw.replace('\\', '/').split('/')[-1].strip()
    name = SAFE_NAME.sub('_', name).lstrip('.')
    return name[:100]


def nice_size(n):
    if n < 1024:
        return '%d bytes' % n
    if n < 1024 * 1024:
        return '%.0f KB' % (n / 1024)
    return '%.1f MB' % (n / 1024 / 1024)


UPDATE_VBS = r'''' Palms LAN Center Setup - installs Palms LAN Center from the Palms Server, or updates it.
' Windows 98: the first time, save palmsget.exe from the Palms Server file page in the same folder.
Option Explicit
Const SERVER = "@@SERVER@@"

Dim fso, sh, list, lines, i, name, got, sys32, lnk, ocxOK, auto, startArgs, startStyle, pg, FOLDER

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
FOLDER = InstallFolder()
' palmsget.exe downloads on every Windows from 98 on. Without it (a new PC), Windows' own parts
' are used: MSXML and MDAC 2.5, which XP has and Windows 98 hasn't.
pg = FindGetter()
auto = HasArg("/auto")
startArgs = ""
startStyle = 1
If auto Then
    startArgs = "/updated"
    If HasArg("/signin") Then startArgs = startArgs & " /signin"
    If HasArg("/min") Then
        ' someone is playing a game: start minimized, and don't take the focus
        startArgs = startArgs & " /min"
        startStyle = 7
    End If
End If

' 1. The program file can't be replaced while it's running
If IsRunning() Then
    If auto Then
        ' Palms LAN Center closes itself after starting this script: give it a few seconds
        For i = 1 To 20
            Nap 500
            If Not IsRunning() Then Exit For
        Next
        If IsRunning() Then
            StopIt
            Nap 1500
        End If
    Else
        If MsgBox("Palms LAN Center is running. Close it and continue?", vbOKCancel + vbQuestion, "Palms LAN Center Setup") <> vbOK Then WScript.Quit
        StopIt
        Nap 1500
        If IsRunning() Then Fail "Palms LAN Center is still running. Close it, then run palms-update.vbs again."
    End If
End If

' 2. Ask the server what to download
list = GetText(SERVER & "/client/manifest.txt")
If Len(list) = 0 Then
    If Len(pg) = 0 Then Fail "Unable to connect to the Palms Server at " & SERVER & "." & vbCrLf & "Check the network connection." & vbCrLf & vbCrLf & _
        "On Windows 98, save palmsget.exe from the Palms Server file page (" & SERVER & "/) in the same folder as palms-update.vbs, then run it again."
    Fail "Unable to connect to the Palms Server at " & SERVER & "." & vbCrLf & "Check the network connection."
End If
If Not fso.FolderExists(FOLDER) Then fso.CreateFolder FOLDER

' 3. Download each file
lines = Split(Replace(list, vbCr, ""), vbLf)
got = 0
For i = 0 To UBound(lines)
    name = Trim(lines(i))
    If Len(name) > 0 Then
        If Not GetFile(SERVER & "/client/" & name, FOLDER & "\" & name) Then Fail "Unable to download " & name & ". Try again later."
        got = got + 1
    End If
Next
' the test PC cleanup script (not required to run Palms LAN Center)
GetFile SERVER & "/palms-clean.vbs", FOLDER & "\palms-clean.vbs"
If Not fso.FileExists(FOLDER & "\PalmsLAN.exe") Then Fail "PalmsLAN.exe is not available on the server. Upload it on the file page from the development PC first."

' files older versions used
On Error Resume Next
For Each name In Array("CartridgeMaker.hta", "pakmaker.gif", "palm.ico", "wget.exe")
    If InStr(1, vbLf & Replace(list, vbCr, "") & vbLf, vbLf & name & vbLf, vbTextCompare) = 0 Then
        If fso.FileExists(FOLDER & "\" & name) Then fso.DeleteFile FOLDER & "\" & name, True
    End If
Next
On Error GoTo 0

' 4. Register the network part (MSWINSCK.OCX) if the server has it. When the LAN Center
'    updates itself it was running a moment ago, so it is registered already.
On Error Resume Next
sys32 = fso.GetSpecialFolder(1).Path
ocxOK = True
If fso.FileExists(FOLDER & "\MSWINSCK.OCX") And Not (auto And fso.FileExists(sys32 & "\MSWINSCK.OCX")) Then
    If Not fso.FileExists(sys32 & "\MSWINSCK.OCX") Then fso.CopyFile FOLDER & "\MSWINSCK.OCX", sys32 & "\"
    ocxOK = False
    If fso.FileExists(sys32 & "\MSWINSCK.OCX") Then ocxOK = (sh.Run("regsvr32 /s """ & sys32 & "\MSWINSCK.OCX""", 0, True) = 0)
End If

' 5. Let it through the Windows XP firewall (does nothing on Windows 98)
sh.Run "netsh firewall add allowedprogram program=""" & FOLDER & "\PalmsLAN.exe"" name=""Palms LAN Center"" mode=ENABLE", 0, True

' 6. Desktop shortcut (Palms LAN Center Setup makes its own, if it was asked to)
If Not fso.FileExists(FOLDER & "\uninst.exe") Then
    Set lnk = sh.CreateShortcut(sh.SpecialFolders("Desktop") & "\Palms LAN Center.lnk")
    lnk.TargetPath = FOLDER & "\PalmsLAN.exe"
    lnk.WorkingDirectory = FOLDER
    lnk.Description = "Palms LAN Center"
    lnk.Save
End If
On Error GoTo 0

' 7. Start it
If Not ocxOK Then
    MsgBox "Palms LAN Center was updated, but MSWINSCK.OCX could not be registered." & vbCrLf & vbCrLf & _
           "Log on to Windows as an administrator and run palms-update.vbs again, or register MSWINSCK.OCX manually.", vbExclamation, "Palms LAN Center Setup"
End If
StartIt
If ocxOK And Not auto Then sh.Popup "Palms LAN Center updated (" & got & " files).", 3, "Palms LAN Center Setup", vbInformation

' Something went wrong. When the LAN Center is updating itself, the copy already installed starts again.
Sub Fail(msg)
    If auto Then
        If fso.FileExists(FOLDER & "\PalmsLAN.exe") Then
            startArgs = Replace(startArgs, "/updated", "/updatefailed")
            StartIt
        End If
    Else
        MsgBox msg, vbExclamation, "Palms LAN Center Setup"
    End If
    WScript.Quit
End Sub

Sub StartIt()
    sh.Run """" & FOLDER & "\PalmsLAN.exe"" " & startArgs, startStyle, False
End Sub

Function HasArg(a)
    Dim k
    HasArg = False
    For k = 0 To WScript.Arguments.Count - 1
        If LCase(WScript.Arguments(k)) = a Then HasArg = True
    Next
End Function

' Where Palms LAN Center is: when it's updating itself, the folder it saved this script in. Otherwise
' the folder it last ran from (program_folder in the registry), or C:\PalmsLAN.
Function InstallFolder()
    Dim f, b
    InstallFolder = "C:\PalmsLAN"
    If HasArg("/auto") Then
        f = fso.GetParentFolderName(WScript.ScriptFullName)
        If fso.FileExists(f & "\PalmsLAN.exe") Then
            InstallFolder = f
            Exit Function
        End If
    End If
    For Each b In Array("HKLM\Software\Palms Arcade", "HKLM\Software\Wow6432Node\Palms Arcade")
        f = RegValue(b & "\LAN Center\program_folder")
        If Len(f) > 0 Then
            If fso.FolderExists(f) Then
                InstallFolder = f
                Exit Function
            End If
        End If
    Next
End Function

' A registry value, or "" when it isn't there
Function RegValue(path)
    On Error Resume Next
    RegValue = ""
    RegValue = CStr(sh.RegRead(path))
    If Err.Number <> 0 Then RegValue = ""
End Function

' palmsget.exe, copied to the temporary folder so it can update itself. "" if there's none.
Function FindGetter()
    Dim src, dst
    FindGetter = ""
    src = FOLDER & "\palmsget.exe"
    If Not fso.FileExists(src) Then src = fso.GetParentFolderName(WScript.ScriptFullName) & "\palmsget.exe"
    If Not fso.FileExists(src) Then Exit Function
    dst = fso.GetSpecialFolder(2).Path & "\palmsget-setup.exe"
    On Error Resume Next
    fso.CopyFile src, dst, True
    If fso.FileExists(dst) Then FindGetter = dst
End Function

Function Cachebust(url)
    Cachebust = url & "?t=" & Replace(CStr(Timer), ",", ".")
End Function

Function GetText(url)
    Dim x, tmp, f
    On Error Resume Next
    GetText = ""
    If Len(pg) > 0 Then
        tmp = fso.GetSpecialFolder(2).Path & "\palms-setup.txt"
        If GetFile(url, tmp) Then
            Set f = fso.OpenTextFile(tmp, 1)
            GetText = f.ReadAll
            f.Close
            fso.DeleteFile tmp
        End If
        Exit Function
    End If
    Set x = CreateObject("MSXML2.XMLHTTP")
    x.Open "GET", Cachebust(url), False
    x.Send
    If Err.Number = 0 Then
        If x.Status = 200 Then GetText = x.responseText
    End If
End Function

Function GetFile(url, path)
    Dim x, st, rc
    On Error Resume Next
    GetFile = False
    If Len(pg) > 0 Then
        ' into a .new file first, so a broken download never replaces a good file
        If fso.FileExists(path & ".new") Then fso.DeleteFile path & ".new", True
        rc = sh.Run("""" & pg & """ -t 3 -T 30 -o """ & fso.GetSpecialFolder(2).Path & "\palms-setup.log"" -O """ & path & ".new"" """ & Cachebust(url) & """", 0, True)
        If rc <> 0 Or Not fso.FileExists(path & ".new") Then Exit Function
        Err.Clear
        If fso.FileExists(path) Then fso.DeleteFile path, True
        fso.MoveFile path & ".new", path
        GetFile = (Err.Number = 0)
        Exit Function
    End If
    Set x = CreateObject("MSXML2.XMLHTTP")
    x.Open "GET", Cachebust(url), False
    x.Send
    If Err.Number <> 0 Then Exit Function
    If x.Status <> 200 Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write x.responseBody
    st.SaveToFile path, 2
    st.Close
    GetFile = (Err.Number = 0)
End Function

Function IsRunning()
    Dim wmi, procs
    On Error Resume Next
    IsRunning = False
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'PalmsLAN.exe'")
    If Err.Number = 0 Then
        IsRunning = (procs.Count > 0)
    Else
        ' Windows 98 has no WMI: a running program's file can't be opened for writing
        IsRunning = InUse(FOLDER & "\PalmsLAN.exe")
    End If
End Function

Function InUse(path)
    Dim f
    InUse = False
    If Not fso.FileExists(path) Then Exit Function
    On Error Resume Next
    Set f = fso.OpenTextFile(path, 8, False)
    If Err.Number <> 0 Then
        InUse = True
    Else
        f.Close
    End If
End Function

' Waits. Windows 98's Script Host 1.0 has no WScript.Sleep: ping waits about a second a try.
Sub Nap(ms)
    On Error Resume Next
    WScript.Sleep ms
    If Err.Number <> 0 Then
        Err.Clear
        sh.Run "ping -n " & (Int(ms / 1000) + 2) & " 127.0.0.1", 0, True
    End If
End Sub

Sub StopIt()
    Dim wmi, p, exe
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process")
        ' the LAN Center, and the Game Library's helpers if they're busy installing a game
        exe = LCase(p.ExecutablePath)
        If LCase(p.Name) = "palmslan.exe" Then
            p.Terminate
        ElseIf (LCase(p.Name) = "palmsget.exe" Or LCase(p.Name) = "wget.exe" Or LCase(p.Name) = "7za.exe") And InStr(exe, LCase(FOLDER) & "\") = 1 Then
            p.Terminate
        End If
    Next
End Sub
'''

CLEAN_VBS = r'''' Palms LAN Center Cleanup - resets a PC: removes every game Palms LAN Center installed, and its settings.
' /quiet: asks nothing (the uninstaller runs it like this).
Option Explicit
Const FOLDER = "C:\PalmsLAN"

Dim fso, sh, here, games, root, ids, i, id, sec, place, keys, k, regText, removed, kept, msg, t, progs, lnks, menu, base, quiet

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

' the LAN Center's folder: where this script is, if PalmsLAN.exe is there too
here = fso.GetParentFolderName(WScript.ScriptFullName)
If Not fso.FileExists(here & "\PalmsLAN.exe") Then here = FOLDER

quiet = (WScript.Arguments.Count > 0)
If quiet Then quiet = (LCase(WScript.Arguments(0)) = "/quiet")
If Not quiet Then
    If MsgBox("This removes all games installed by Palms LAN Center on this PC, including their registry " & _
              "settings and Windows Firewall exceptions, any incomplete downloads, and the Palms LAN Center settings." & vbCrLf & vbCrLf & _
              "Configuration files and saved games in the game folders are also deleted. The Palms LAN Center program is not removed." & vbCrLf & vbCrLf & _
              "Continue?", vbOKCancel + vbExclamation + vbDefaultButton2, "Palms LAN Center Cleanup") <> vbOK Then WScript.Quit
End If

' 1. Close the LAN Center, so nothing is in use
If IsRunning() Then
    StopIt
    Nap 1500
    If IsRunning() And Not quiet Then
        MsgBox "Close Palms LAN Center first, then run palms-clean.vbs again.", vbExclamation, "Palms LAN Center Cleanup"
        WScript.Quit
    End If
End If

' 2. Each game it installed: from the registry, and from gamepaks.ini (older versions)
base = RegBase()
Set games = ReadIni(here & "\gamepaks.ini")
If Len(base) > 0 Then AddRegGames games
root = Setting("games_folder")
If Len(root) = 0 Then root = "C:\LANPartyGames"
removed = ""
kept = ""
regText = ""
If games.Exists("installed") Then
    ids = Split(IniValue(games, "installed", "list"), ",")
    For i = 0 To UBound(ids)
        id = LCase(Trim(ids(i)))
        If Len(id) > 0 Then regText = regText & "[-HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\PalmsArcade_" & id & "]" & vbCrLf & vbCrLf
        If Len(id) > 0 And games.Exists(id) Then
            Set sec = games(id)
            place = Get1(sec, "folder")
            ' the firewall exceptions for its programs (a pack can have several). No netsh on Windows 98.
            progs = Get1(sec, "programs")
            If Len(progs) = 0 Then progs = Get1(sec, "exe")
            progs = Split(progs, ";")
            For k = 0 To UBound(progs)
                If Len(Trim(progs(k))) > 0 And Len(place) > 0 Then
                    On Error Resume Next
                    sh.Run "netsh firewall delete allowedprogram program=""" & place & "\" & Trim(progs(k)) & """", 0, True
                    On Error GoTo 0
                End If
            Next
            ' its Start menu shortcuts
            lnks = Split(Get1(sec, "shortcuts"), ";")
            For k = 0 To UBound(lnks)
                If LCase(Right(Trim(lnks(k)), 4)) = ".lnk" Then
                    On Error Resume Next
                    If fso.FileExists(Trim(lnks(k))) Then fso.DeleteFile Trim(lnks(k)), True
                    On Error GoTo 0
                End If
            Next
            ' its registry keys, removed all at once below
            keys = Split(Get1(sec, "regkeys"), ";")
            For k = 0 To UBound(keys)
                If Len(Trim(keys(k))) > 0 Then regText = regText & "[-" & Trim(keys(k)) & "]" & vbCrLf & vbCrLf
            Next
            ' its folder
            If SafeFolder(place) Then
                On Error Resume Next
                If fso.FolderExists(place) Then fso.DeleteFolder place, True
                On Error GoTo 0
                If fso.FolderExists(place) Then
                    kept = kept & "    " & place & vbCrLf
                Else
                    removed = removed & "    " & Get1(sec, "name") & vbCrLf
                End If
            End If
        End If
    Next
End If
' and the LAN Center's own settings and list of games
regText = regText & "[-HKEY_LOCAL_MACHINE\Software\Palms Arcade]" & vbCrLf & vbCrLf
Set t = fso.CreateTextFile(fso.GetSpecialFolder(2).Path & "\palms-clean.reg", True, False)
t.Write "REGEDIT4" & vbCrLf & vbCrLf & regText
t.Close
' Palms LAN Center and 32-bit games keep their settings on the 32-bit side: on 64-bit Windows that's SysWOW64's regedit
t = fso.GetSpecialFolder(0).Path & "\SysWOW64\regedit.exe"
If Not fso.FileExists(t) Then t = "regedit"
sh.Run """" & t & """ /s """ & fso.GetSpecialFolder(2).Path & "\palms-clean.reg""", 0, True
On Error Resume Next
fso.DeleteFile fso.GetSpecialFolder(2).Path & "\palms-clean.reg"
On Error GoTo 0

' 3. Half-finished downloads, and the games folder itself if nothing else is in it
On Error Resume Next
If fso.FolderExists(root & "\_downloads") Then fso.DeleteFolder root & "\_downloads", True
If fso.FolderExists(root) Then
    If fso.GetFolder(root).Files.Count = 0 And fso.GetFolder(root).SubFolders.Count = 0 Then fso.DeleteFolder root, True
End If

' The Start menu folder the shortcuts were in, if it's empty now
menu = sh.SpecialFolders("Programs") & "\Palms Arcade Games"
If fso.FolderExists(menu) Then
    If fso.GetFolder(menu).Files.Count = 0 And fso.GetFolder(menu).SubFolders.Count = 0 Then fso.DeleteFolder menu, True
End If

' 4. Settings files from older versions, cached game icons and work files
For Each t In Array("palms.ini", "gamepaks.ini", "palms.ini.old", "gamepaks.ini.old")
    If fso.FileExists(here & "\" & t) Then fso.DeleteFile here & "\" & t, True
Next
If fso.FolderExists(here & "\icons") Then fso.DeleteFolder here & "\icons", True
If fso.FolderExists(here & "\temp") Then fso.DeleteFolder here & "\temp", True
On Error GoTo 0

msg = "Cleanup complete. The Palms Server will treat this PC as a new PC." & vbCrLf & vbCrLf
If Len(removed) > 0 Then
    msg = msg & "Games removed:" & vbCrLf & removed & vbCrLf
Else
    msg = msg & "No games were installed." & vbCrLf & vbCrLf
End If
If Len(kept) > 0 Then msg = msg & "The following folders could not be removed completely. A game may still be running:" & vbCrLf & kept & vbCrLf
If Not quiet Then MsgBox msg, vbInformation, "Palms LAN Center Cleanup"

' ---------------------------------------------------------------- the registry
' Where Palms LAN Center keeps its settings, or "" if it hasn't any. It is a 32-bit program:
' on 64-bit Windows its HKEY_LOCAL_MACHINE\Software is Software\Wow6432Node.
Function RegBase()
    Dim b
    RegBase = ""
    For Each b In Array("HKLM\Software\Palms Arcade", "HKLM\Software\Wow6432Node\Palms Arcade")
        If Len(RegValue(b & "\LAN Center\pcid")) > 0 Or Len(RegValue(b & "\Games\Installed")) > 0 Then
            RegBase = b
            Exit Function
        End If
    Next
End Function

' A value, or "" when it isn't there
Function RegValue(path)
    On Error Resume Next
    RegValue = ""
    RegValue = CStr(sh.RegRead(path))
    If Err.Number <> 0 Then RegValue = ""
End Function

' The installed games in the registry, added to the ones from gamepaks.ini
Sub AddRegGames(d)
    Dim list, ids2, j, gid, inner, key, both
    list = RegValue(base & "\Games\Installed")
    If Len(list) = 0 Then Exit Sub
    both = IniValue(d, "installed", "list")
    If Len(both) > 0 Then both = both & ","
    If Not d.Exists("installed") Then d.Add "installed", NewDict()
    Set inner = d("installed")
    inner("list") = both & list
    ids2 = Split(list, ",")
    For j = 0 To UBound(ids2)
        gid = LCase(Trim(ids2(j)))
        If Len(gid) > 0 And Not d.Exists(gid) Then
            Set inner = NewDict()
            For Each key In Array("name", "folder", "programs", "exe", "shortcuts", "regkeys")
                inner(key) = RegValue(base & "\Games\" & gid & "\" & key)
            Next
            d.Add gid, inner
        End If
    Next
End Sub

' A LAN Center setting: the registry, or palms.ini (older versions)
Function Setting(key)
    Setting = ""
    If Len(base) > 0 Then Setting = RegValue(base & "\LAN Center\" & key)
    If Len(Setting) = 0 Then Setting = IniValue(ReadIni(here & "\palms.ini"), "client", key)
End Function

' ---------------------------------------------------------------- the old .ini files
Function NewDict()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set NewDict = d
End Function

' An .ini file as a dictionary of sections, each a dictionary of key = value
Function ReadIni(path)
    Dim d, f, lines, i, line, cur, p, inner
    Set d = NewDict()
    Set ReadIni = d
    If Not fso.FileExists(path) Then Exit Function
    If fso.GetFile(path).Size = 0 Then Exit Function
    Set f = fso.OpenTextFile(path, 1, False, 0)
    lines = Split(Replace(f.ReadAll, vbCr, ""), vbLf)
    f.Close
    cur = ""
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            cur = LCase(Mid(line, 2, Len(line) - 2))
            If Not d.Exists(cur) Then d.Add cur, NewDict()
        ElseIf Len(cur) > 0 Then
            p = InStr(line, "=")
            If p > 1 Then
                Set inner = d(cur)
                inner(LCase(Trim(Left(line, p - 1)))) = Trim(Mid(line, p + 1))
            End If
        End If
    Next
End Function

Function IniValue(d, section, key)
    IniValue = ""
    If d.Exists(section) Then IniValue = Get1(d(section), key)
End Function

Function Get1(sec, key)
    Get1 = ""
    If sec.Exists(key) Then Get1 = sec(key)
End Function

' Never a drive, the Windows folder or Program Files, whatever the list of games says
Function SafeFolder(p)
    Dim low
    SafeFolder = False
    low = LCase(p)
    If Len(p) < 4 Or Mid(p, 2, 2) <> ":\" Or InStr(p, "..") > 0 Then Exit Function
    If low = LCase(fso.GetSpecialFolder(0).Path) Or low = LCase(sh.ExpandEnvironmentStrings("%ProgramFiles%")) Then Exit Function
    If low = LCase(here) Or low = LCase(root) Then Exit Function
    SafeFolder = True
End Function

' ---------------------------------------------------------------- the running program
Function IsRunning()
    Dim wmi, p
    On Error Resume Next
    IsRunning = False
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then
        ' Windows 98 has no WMI: a running program's file can't be opened for writing
        Err.Clear
        IsRunning = InUse(here & "\PalmsLAN.exe")
        Exit Function
    End If
    For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process")
        If LCase(p.Name) = "palmslan.exe" Then IsRunning = True
    Next
End Function

Function InUse(path)
    Dim f
    InUse = False
    If Not fso.FileExists(path) Then Exit Function
    On Error Resume Next
    Set f = fso.OpenTextFile(path, 8, False)
    If Err.Number <> 0 Then
        InUse = True
    Else
        f.Close
    End If
End Function

Sub StopIt()
    Dim wmi, p
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If Err.Number <> 0 Then Exit Sub
    For Each p In wmi.ExecQuery("SELECT * FROM Win32_Process")
        Select Case LCase(p.Name)
        Case "palmslan.exe", "palmsget.exe", "wget.exe", "7za.exe"
            p.Terminate
        End Select
    Next
End Sub

' Waits. Windows 98's Script Host 1.0 has no WScript.Sleep: ping waits about a second a try.
Sub Nap(ms)
    On Error Resume Next
    WScript.Sleep ms
    If Err.Number <> 0 Then
        Err.Clear
        sh.Run "ping -n " & (Int(ms / 1000) + 2) & " 127.0.0.1", 0, True
    End If
End Sub
'''


def make_files_handler(root, server_name):
    client_dir = root / 'client'

    class FilesHandler(http.server.BaseHTTPRequestHandler):
        server_version = 'PalmsFiles/' + VERSION
        protocol_version = 'HTTP/1.0'

        def log_message(self, fmt, *args):
            pass                                  # quiet; uploads are logged below

        # ----- helpers
        def base_url(self):
            host = self.headers.get('Host') or ('%s:%d' % self.server.server_address)
            return 'http://' + host

        def send_bytes(self, code, data, ctype, extra=None):
            self.send_response(code)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Pragma', 'no-cache')
            self.send_header('Expires', '0')
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            if self.command != 'HEAD':
                self.wfile.write(data)

        def resolve(self, url_path):
            """A real file under the files folder, or None. Never anything outside it."""
            rel = urllib.parse.unquote(url_path.split('?', 1)[0]).lstrip('/')
            try:
                p = (root / rel).resolve()
                p.relative_to(root.resolve())
            except (ValueError, OSError):
                return None
            return p if p.is_file() else None

        def save(self, folder, raw_name, data, who):
            name = safe_name(raw_name)
            if not name:
                return None, 'Invalid file name.'
            folder.mkdir(parents=True, exist_ok=True)
            tmp = folder / ('.%s.part' % name)
            try:
                tmp.write_bytes(data)
                os.replace(tmp, folder / name)
            except OSError as e:
                return None, ("Unable to save files in %s (%s). On the server, run: chown -R palms:palms %s"
                              % (folder, e.strerror, root))
            log('File page: %s uploaded %s (%s) from %s.' % (who, folder.name + '/' + name, nice_size(len(data)), self.client_address[0]))
            return name, None

        def manifest(self):
            names = sorted(p.name for p in client_dir.glob('*') if p.is_file() and not p.name.startswith('.'))
            names = [n for n in names if n not in CLIENT_LAST] + [n for n in CLIENT_LAST if n in names]
            return ('\r\n'.join(names) + '\r\n').encode()

        # ----- GET and HEAD
        def do_HEAD(self):
            self.do_GET()

        def do_GET(self):
            path = self.path.split('?', 1)[0]
            if path in ('/', '/index.html'):
                return self.send_bytes(200, self.index_page().encode('utf-8'), 'text/html; charset=utf-8')
            if path == '/client/manifest.txt':
                return self.send_bytes(200, self.manifest(), 'text/plain')
            if path in ('/palms-update.vbs', '/palms-clean.vbs'):
                body = {'/palms-update.vbs': UPDATE_VBS, '/palms-clean.vbs': CLEAN_VBS}[path].replace('@@SERVER@@', self.base_url())
                return self.send_bytes(200, body.replace('\n', '\r\n').encode('cp1252'), 'application/octet-stream',
                                       {'Content-Disposition': 'attachment; filename="%s"' % path[1:]})
            p = self.resolve(path)
            if not p:
                return self.send_bytes(404, b'File not found.', 'text/plain')
            ctype = mimetypes.guess_type(p.name)[0] or 'application/octet-stream'
            extra = {}
            if p.suffix.lower() in DOWNLOAD_ONLY:
                ctype = 'application/octet-stream'
                extra['Content-Disposition'] = 'attachment; filename="%s"' % p.name
            self.send_bytes(200, p.read_bytes(), ctype, extra)

        # ----- POST /upload from the form on the page
        def do_POST(self):
            if self.path.split('?', 1)[0] != '/upload':
                return self.send_bytes(404, b'File not found.', 'text/plain')
            length = int(self.headers.get('Content-Length') or 0)
            ctype = self.headers.get('Content-Type', '')
            m = re.search(r'boundary=("?)([^";]+)\1', ctype)
            if not m or length <= 0:
                return self.page_with('Select a file to upload.', bad=True)
            if length > MAX_UPLOAD:
                self.rfile.read(length)
                return self.page_with('File too large. Maximum file size: %s.' % nice_size(MAX_UPLOAD), bad=True)
            body = self.rfile.read(length)
            fields, filename, data = {}, '', None
            for part in body.split(b'--' + m.group(2).encode()):
                head, sep, content = part.partition(b'\r\n\r\n')
                if not sep:
                    continue
                if content.endswith(b'\r\n'):
                    content = content[:-2]
                disp = head.decode('latin-1')
                fname = re.search(r'name="([^"]*)"', disp)
                fn = re.search(r'filename="([^"]*)"', disp)
                if fn is not None:
                    filename, data = fn.group(1), content
                elif fname:
                    fields[fname.group(1)] = content.decode('utf-8', 'replace')
            if not filename or not data:
                return self.page_with('Select a file to upload.', bad=True)
            # the LAN Center's own files go where palms-update.vbs and the updates find them, named exactly
            known = {n.lower(): n for n in CLIENT_FILES}
            folder, old = root, ''
            if safe_name(filename).lower() in known:
                folder, filename = client_dir, known[safe_name(filename).lower()]
                if filename == 'PalmsLAN.exe':
                    old = exe_version(client_dir / filename)
            name, err = self.save(folder, filename, data, 'the upload form')
            if err:
                return self.page_with(err, bad=True)
            if folder is not client_dir:
                return self.page_with('Uploaded %s to Other Files.' % name)
            if name != 'PalmsLAN.exe':
                return self.page_with('Uploaded %s to LAN Center Files. palms-update.vbs installs it on each PC.' % name)
            new = exe_version(client_dir / name)
            if not new:
                return self.page_with('Uploaded PalmsLAN.exe to LAN Center Files, but its version could not be read. '
                                      'PCs do not update to it automatically.', bad=True)
            if old and ver_tuple(new) == ver_tuple(old):
                return self.page_with('Uploaded PalmsLAN.exe %s to LAN Center Files. This is the same version as before, so PCs '
                                      'do not update. In Visual Basic, raise the version (Project > Properties > Make), '
                                      'then make PalmsLAN.exe and upload it again.' % new, bad=True)
            if old and ver_tuple(new) < ver_tuple(old):
                return self.page_with('Uploaded PalmsLAN.exe %s to LAN Center Files. It is older than the version it replaced '
                                      '(%s), so PCs do not update to it.' % (new, old), bad=True)
            self.page_with('Published Palms LAN Center %s. PCs with 0.7.6 or later update themselves within a minute. '
                           'Run palms-update.vbs on new PCs and on PCs with an earlier version.' % new)

        def page_with(self, msg, bad=False):
            self.send_bytes(200, self.index_page(msg, bad).encode('utf-8'), 'text/html; charset=utf-8')

        # ----- the page itself: plain HTML that Internet Explorer 6 is happy with
        def index_page(self, msg='', bad=False):
            def rows(folder, prefix):
                items = sorted((p for p in folder.glob('*') if p.is_file() and not p.name.startswith('.')),
                               key=lambda p: p.name.lower())
                if not items:
                    return '<tr><td colspan="3" class="none">No files.</td></tr>'
                out = []
                for p in items:
                    st = p.stat()
                    out.append('<tr><td><a href="%s%s">%s</a></td><td class="r">%s</td><td>%s</td></tr>' % (
                        prefix, urllib.parse.quote(p.name), _html.escape(p.name), nice_size(st.st_size),
                        time.strftime('%a %d %b, %I:%M %p', time.localtime(st.st_mtime)).replace(' 0', ' ')))
                return ''.join(out)
            note = ''
            if msg:
                note = '<p class="%s">%s</p>' % ('bad' if bad else 'good', _html.escape(msg))
            return PAGE % {
                'name': _html.escape(server_name), 'note': note,
                'client_rows': rows(client_dir, '/client/'), 'other_rows': rows(root, '/'),
                'limit': nice_size(MAX_UPLOAD)}

    return FilesHandler


PAGE = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Palms Server - Files</title>
<style type="text/css">
body { background: #008284; margin: 0; padding: 16px; font: 13px Tahoma, "MS Sans Serif", Arial, sans-serif; color: #000; }
.win { background: #C6C3C6; border: 2px outset #FFFFFF; max-width: 760px; margin: 0 auto; }
.title { background: #08599C; color: #FFF; font-weight: bold; padding: 4px 8px; }
.body { padding: 12px 14px; }
h2 { font-size: 15px; color: #103473; margin: 16px 0 6px; }
.box { background: #FFFFE7; border: 1px solid #000; padding: 8px 10px; margin: 8px 0; }
table { width: 100%%; border-collapse: collapse; background: #FFF; border: 2px inset #FFF; }
th { text-align: left; background: #DEDFDE; padding: 3px 6px; font-weight: normal; border-bottom: 1px solid #848284; }
td { padding: 3px 6px; border-bottom: 1px solid #EEE; }
td.r { text-align: right; white-space: nowrap; }
td.none { color: #636163; }
.good { background: #FFF; border: 1px solid #188221; color: #188221; padding: 6px 8px; font-weight: bold; }
.bad { background: #FFF; border: 1px solid #AD1C10; color: #AD1C10; padding: 6px 8px; font-weight: bold; }
.small { color: #424142; font-size: 12px; }
</style></head>
<body><div class="win"><div class="title">Palms Server Files - %(name)s</div><div class="body">
%(note)s
<div class="box"><b>Install or update Palms LAN Center:</b> save
<a href="/palms-update.vbs">palms-update.vbs</a> to the desktop and double-click it.
The latest version is installed to C:\\PalmsLAN and started. Run it again to install a new version.
<br><b>Windows 98:</b> the first time, also save <a href="/client/palmsget.exe">palmsget.exe</a> in the same folder as palms-update.vbs.</div>
<div class="box"><b>Publish a new version (development PC):</b> after File &gt; Make PalmsLAN.exe in Visual Basic 6.0,
upload PalmsLAN.exe with the form at the bottom of this page. PCs with 0.7.6 or later update themselves within a minute.</div>
<div class="box"><b>Test PC cleanup:</b> <a href="/palms-clean.vbs">palms-clean.vbs</a> removes all games installed by
Palms LAN Center, and its settings. palms-update.vbs also copies it to C:\\PalmsLAN.</div>

<h2>LAN Center Files</h2>
<table><tr><th>Name</th><th>Size</th><th>Modified</th></tr>%(client_rows)s</table>

<h2>Other Files</h2>
<table><tr><th>Name</th><th>Size</th><th>Modified</th></tr>%(other_rows)s</table>

<h2>Upload File</h2>
<form method="post" action="/upload" enctype="multipart/form-data">
<table><tr><td>File:</td><td><input type="file" name="file" size="40"></td></tr>
<tr><td></td><td><input type="submit" value="Upload"> <span class="small">Maximum file size: %(limit)s. A file with the same name is replaced.</span></td></tr>
<tr><td></td><td class="small">PalmsLAN.exe, MSWINSCK.OCX, palmsget.exe, 7za.exe and the four .wav sounds go to LAN Center Files.
Everything else goes to Other Files.</td></tr></table>
</form>
</div></div></body></html>
'''


# ---------------------------------------------------------------- server

class Client:
    def __init__(self, server, reader, writer):
        self.server = server
        self.reader = reader
        self.writer = writer
        peer = writer.get_extra_info('peername') or ('?', 0)
        self.addr = peer[0]
        self.pcid = None          # set by HELLO; telnet users never send one
        self.ver = ''             # LAN Center version, from HELLO
        self.pcname = ''          # Windows computer name, from HELLO
        self.specs = [''] * 8     # processor, memory, graphics, Windows, screen, sound, network card, link Mb (SPECS)
        self.caps = set()         # what the client can do, from CAPS ('ping' = answers SPING)
        self.ms = None            # its ping from the server, in milliseconds
        self.nick = None
        self.seat = None
        self.admin = False
        self.status = 'here'
        self.closed = False

    @property
    def signed_in(self):
        return self.nick is not None

    def send(self, *fields):
        if self.closed:
            return
        line = '|'.join(str(f) for f in fields) + '\r\n'
        try:
            self.writer.write(line.encode('cp1252', 'replace'))
        except Exception:
            self.closed = True

    def label(self):
        return self.nick or self.addr


class PalmsServer:
    def __init__(self, cfg):
        self.cfg = cfg
        self.name = cfg['name']
        self.tcp_port = int(cfg['tcp_port'])
        self.beacon_port = int(cfg['beacon_port'])
        self.beacon_seconds = max(1, int(cfg['beacon_seconds']))
        self.admin_password = cfg['admin_password']
        self.clients = set()
        self.history = collections.deque(maxlen=HISTORY_LINES)
        self.file_page = False    # the file page is running, so LAN Centers can update from it
        self.client_ver = ''      # the version of files/client/PalmsLAN.exe
        self._client_stat = None
        self._announced = ''
        self.updating = {}        # nick (lower case) -> when they closed to install an update
        DATA_DIR.mkdir(exist_ok=True)
        self.pcs = self._load_pcs()
        self.schedule = self._load_schedule()
        self.templates = self._load_templates()
        self.vote = None          # the current or last vote, see start_vote()
        self.suggest = []         # games suggested for the next vote: {'name', 'by': [nick, ...]}
        self.food = self._load_food()
        self.cups = self._load_cups()
        self.arrived = {}         # nick (lower case) -> when they first signed in today
        self.gone = {}            # nick (lower case) -> who left, when and why (for Who's Here)
        self.maps, self.map_active = self._load_maps()
        self.paks = []            # PAK and GAME lines from the pak server's catalog.txt, passed on as they are
        self.paks_base = ''       # the pak server's address, for downloading, like http://10.0.200.73/
        self.paks_note = 'Retrieving the game list from the Cartridge Server...'   # why the list is empty, if it is

    # ----- remembered PCs
    def _load_pcs(self):
        try:
            return json.loads(PCS_FILE.read_text(encoding='utf-8'))
        except FileNotFoundError:
            return {}
        except Exception as e:
            log('Could not read %s (%s). Starting with no remembered PCs.' % (PCS_FILE, e))
            return {}

    def _save_pcs(self):
        tmp = PCS_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.pcs, indent=1), encoding='utf-8')
        os.replace(tmp, PCS_FILE)

    # ----- today's schedule
    def _load_schedule(self):
        try:
            data = json.loads(SCHED_FILE.read_text(encoding='utf-8'))
            return [e for e in data if TIME_RE.match(e.get('time', ''))]
        except FileNotFoundError:
            return []
        except Exception as e:
            log('Could not read %s (%s). Starting with an empty schedule.' % (SCHED_FILE, e))
            return []

    def _save_schedule(self):
        tmp = SCHED_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.schedule, indent=1), encoding='utf-8')
        os.replace(tmp, SCHED_FILE)

    def send_schedule(self, c):
        events = sorted(self.schedule, key=lambda e: (party_minutes(e['time']), e['id']))
        for e in events:
            c.send('SCHED', e['id'], e['time'], e['icon'], e['name'])
        # sent separately, so older LAN Centers (which ignore it) still read SCHED correctly
        for e in events:
            ins = e.get('in', [])
            c.send('SCHEDINFO', e['id'], len(ins), ','.join(ins), e.get('note', ''))
        c.send('SCHEDEND', len(events))

    def schedule_changed(self):
        self._save_schedule()
        for c in self.signed_in_clients():
            self.send_schedule(c)

    # ----- schedule templates
    def _load_templates(self):
        try:
            return json.loads(TPL_FILE.read_text(encoding='utf-8'))
        except FileNotFoundError:
            tpls = [{'id': i + 1, 'name': n, 'lines': l} for i, (n, l) in enumerate(DEFAULT_TEMPLATES)]
            try:
                TPL_FILE.write_text(json.dumps(tpls, indent=1), encoding='utf-8')
            except OSError:
                pass
            return tpls
        except Exception as e:
            log('Could not read %s (%s). Using the built-in templates.' % (TPL_FILE, e))
            return [{'id': i + 1, 'name': n, 'lines': l} for i, (n, l) in enumerate(DEFAULT_TEMPLATES)]

    def _save_templates(self):
        tmp = TPL_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.templates, indent=1), encoding='utf-8')
        os.replace(tmp, TPL_FILE)

    def send_template_list(self, c):
        for t in self.templates:
            c.send('TPL', t['id'], t['name'])
        c.send('TPLEND', len(self.templates))

    @staticmethod
    def template_lines(raw):
        """Split a template sent as one line. Tabs are kept, they line the times up nicely."""
        out = []
        for l in raw.split('~'):
            l = ''.join(ch for ch in l.replace('|', '-') if ch == '\t' or (ch >= ' ' and ch != '\x7f')).strip()[:80]
            if l:
                out.append(l)
        return out[:40]

    def apply_template(self, c, mode, lines):
        items = [parse_template_line(l) for l in lines]
        good = [i for i in items if i]
        if not good:
            c.send('ERR', 'NO_TIMES', 'No lines start with a time (for example, 12:15 PM Game 1).')
            return
        if mode == 'replace':
            self.schedule = []
        next_id = max([e['id'] for e in self.schedule] + [0]) + 1
        for n, (t, icon, text) in enumerate(good):
            self.schedule.append({'id': next_id + n, 'time': t, 'icon': icon, 'name': text})
        self.schedule_changed()
        c.send('TPLAPPLIED', len(good), len(items) - len(good))
        if mode == 'replace':
            self.note('calendar', "%s updated today's schedule. See What's On." % c.nick)
        else:
            self.note('calendar', "%s added %d event%s to today's schedule." % (c.nick, len(good), '' if len(good) == 1 else 's'))
        log('%s used a template: %s, %d items' % (c.nick, mode, len(good)))

    # ----- next-game voting
    def vote_counts(self):
        counts = [0] * len(self.vote['options'])
        for idx in self.vote['votes'].values():
            counts[idx] += 1
        return counts

    def send_vote(self, c):
        v = self.vote
        if not v:
            c.send('VOTEINFO', 'none', 0, 0, '')
            c.send('VOTEEND', 0, 0)
            return
        counts = self.vote_counts()
        left = 0
        if v['state'] == 'open':
            left = max(0, int(v['deadline'] - asyncio.get_running_loop().time() + 0.5))
        mine = v['votes'].get(c.nick.lower(), -1) + 1
        c.send('VOTEINFO', v['state'], left, mine, v['question'])
        for i, name in enumerate(v['options']):
            c.send('VOTEOPT', i + 1, counts[i], name)
        c.send('VOTEEND', sum(counts), v.get('winner', 0))

    def vote_changed(self):
        for c in self.signed_in_clients():
            self.send_vote(c)

    # ----- games suggested for the next vote
    def send_suggestions(self, c):
        for s in self.suggest:
            c.send('SUGG', len(s['by']), ','.join(s['by']), s['name'])
        c.send('SUGGEND', len(self.suggest))

    def suggestions_changed(self):
        self.suggest = [s for s in self.suggest if s['by']]
        self.suggest.sort(key=lambda s: (-len(s['by']), s['name'].lower()))
        for c in self.signed_in_clients():
            self.send_suggestions(c)

    def find_suggestion(self, name):
        return next((s for s in self.suggest if s['name'].lower() == name.lower()), None)

    # ----- the food order
    def _load_food(self):
        try:
            f = json.loads(FOOD_FILE.read_text(encoding='utf-8'))
            if f.get('state') in ('open', 'closed') and isinstance(f.get('orders'), list):
                return f
        except FileNotFoundError:
            pass
        except Exception as e:
            log('Could not read %s (%s). Starting with no food order.' % (FOOD_FILE, e))
        return {'state': 'closed', 'title': '', 'orders': []}

    def send_food(self, c):
        f = self.food
        c.send('FOODINFO', f['state'], len(f['orders']), f['title'])
        for o in f['orders']:
            c.send('FOODITEM', o['nick'], o['time'], o['text'])
        c.send('FOODEND', len(f['orders']))

    def food_changed(self):
        tmp = FOOD_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.food, indent=1), encoding='utf-8')
        os.replace(tmp, FOOD_FILE)
        for c in self.signed_in_clients():
            self.send_food(c)

    # ----- tournaments: single elimination. Players sign up, the server draws the bracket,
    # players report their results and the host confirms them.
    def _load_cups(self):
        try:
            data = json.loads(CUPS_FILE.read_text(encoding='utf-8'))
            cups = [t for t in data if t.get('state') in CUP_STATES and isinstance(t.get('players'), list)]
            for t in cups:
                t.setdefault('slots', [])
                t.setdefault('results', {})
                t.setdefault('reports', {})
                t.setdefault('champion', '')
            return cups
        except FileNotFoundError:
            return []
        except Exception as e:
            log('Could not read %s (%s). Starting with no tournaments.' % (CUPS_FILE, e))
            return []

    def _save_cups(self):
        tmp = CUPS_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.cups, indent=1), encoding='utf-8')
        os.replace(tmp, CUPS_FILE)

    def find_cup(self, cup_id):
        return next((t for t in self.cups if str(t['id']) == str(cup_id)), None)

    @staticmethod
    def round_name(r, rounds):
        left = rounds - r
        return {1: 'Final', 2: 'Semi-final', 3: 'Quarter-final'}.get(left, 'Round %d' % (r + 1))

    @staticmethod
    def seed_order(n):
        """Bracket positions for seeds 1..n (n a power of 2), so seeds 1 and 2 can only meet in the final
        and the byes (the highest seeds that have no player) never meet each other."""
        order = [1]
        while len(order) < n:
            size = len(order) * 2
            order = [x for s in order for x in (s, size + 1 - s)]
        return order

    @staticmethod
    def cup_bracket(t):
        """The whole bracket from the draw and the results so far. A list of rounds, each a list of
        matches: a and b are a nickname, None for a bye, or '' while still to be decided; win is
        'a', 'b' or ''; out is who goes through."""
        slots = t['slots']
        rounds, prev, r = [], None, 0
        size = len(slots)
        while size >= 2:
            ms = []
            for i in range(size // 2):
                if r == 0:
                    a, b = slots[2 * i], slots[2 * i + 1]
                else:
                    a, b = prev[2 * i]['out'], prev[2 * i + 1]['out']
                key = '%d.%d' % (r, i)
                m = {'r': r, 'i': i, 'key': key, 'a': a, 'b': b, 'win': '', 'score': '', 'auto': False, 'rep': None}
                if a is None and b:
                    m['win'], m['auto'] = 'b', True
                elif b is None and a:
                    m['win'], m['auto'] = 'a', True
                elif a and b:
                    res = t['results'].get(key)
                    if res and res.get('a') == a and res.get('b') == b and res.get('win') in ('a', 'b'):
                        m['win'], m['score'] = res['win'], res.get('score', '')
                    else:
                        rep = t['reports'].get(key)
                        if rep and rep.get('a') == a and rep.get('b') == b:
                            m['rep'] = rep
                if m['win'] == 'a':
                    m['out'] = a
                elif m['win'] == 'b':
                    m['out'] = b
                elif a is None and b is None:
                    m['out'] = None
                else:
                    m['out'] = ''
                ms.append(m)
            rounds.append(ms)
            prev = ms
            size //= 2
            r += 1
        return rounds

    def cup_tidy(self, t):
        """Forget results and reports that no longer fit the bracket, and settle who won."""
        rounds = self.cup_bracket(t)
        live = {m['key']: m for ms in rounds for m in ms}
        t['results'] = {k: v for k, v in t['results'].items()
                        if k in live and live[k]['win'] and not live[k]['auto'] and live[k]['a'] == v.get('a')}
        t['reports'] = {k: v for k, v in t['reports'].items() if k in live and live[k]['rep'] is v}
        champ = rounds[-1][0]['out'] if rounds else ''
        if t['state'] in ('running', 'done'):
            t['champion'] = champ or ''
            t['state'] = 'done' if champ else 'running'
        return rounds

    @staticmethod
    def ready_matches(rounds):
        return {m['key']: m for ms in rounds for m in ms if m['a'] and m['b'] and not m['win']}

    def send_cup(self, c, t):
        f = lambda x: str(x or '').replace('|', '/')
        wire = lambda p: '-' if p is None else p
        c.send('CUP', t['id'], t['state'], t['max'], f(t['time']), f(t['game']), t.get('champion', ''), t['name'])
        c.send('CUPNOTE', t['id'], f(t['notes']))
        rounds = []
        if t['state'] == 'signup':
            c.send('CUPP', t['id'], ','.join(t['players']))
        else:
            c.send('CUPP', t['id'], ','.join(wire(p) for p in t['slots']))
            rounds = self.cup_bracket(t)
            for ms in rounds:
                for m in ms:
                    rep = m['rep'] or {}
                    c.send('CUPM', t['id'], m['r'], m['i'], wire(m['a']), wire(m['b']), m['win'], f(m['score']),
                           rep.get('by', ''), rep.get('win', ''), f(rep.get('score', '')))
        c.send('CUPEND', t['id'], len(rounds))

    def send_cups(self, c):
        c.send('CUPCLEAR', len(self.cups))
        for t in self.cups:
            self.send_cup(c, t)

    def cup_changed(self, t, before=None):
        """Save, tell everyone, and tell the players whose match has just become ready."""
        rounds = self.cup_tidy(t) if t['state'] != 'signup' else []
        self._save_cups()
        for c in self.signed_in_clients():
            if 'cups' in c.caps:
                self.send_cup(c, t)
        if t['state'] == 'running' and before is not None:
            for key, m in self.ready_matches(rounds).items():
                if '%s|%s|%s' % (key, m['a'], m['b']) in before:
                    continue
                for me, other in ((m['a'], m['b']), (m['b'], m['a'])):
                    p = self.find_nick(me)
                    if p and 'cups' in p.caps:
                        p.send('CUPMATCH', t['id'], m['r'], m['i'], other)

    def cup_ready_now(self, t):
        if t['state'] != 'running':
            return set()
        return {'%s|%s|%s' % (k, m['a'], m['b']) for k, m in self.ready_matches(self.cup_bracket(t)).items()}

    def cup_match(self, t, r, i):
        try:
            r, i = int(r), int(i)
        except ValueError:
            return None, None
        rounds = self.cup_bracket(t)
        if not 0 <= r < len(rounds) or not 0 <= i < len(rounds[r]):
            return None, rounds
        return rounds[r][i], rounds

    def cup_draw(self, t):
        players = list(t['players'])
        random.shuffle(players)
        n = 2
        while n < len(players):
            n *= 2
        slots = [None] * n
        for pos, seed in enumerate(self.seed_order(n)):
            if seed <= len(players):
                slots[pos] = players[seed - 1]
        t['slots'] = slots
        t['results'] = {}
        t['reports'] = {}
        t['champion'] = ''
        t['state'] = 'drawn'

    def cup_command(self, c, cmd, parts, arg):
        """CUP... commands. Returns True when handled."""
        host_only = ('CUPNEW', 'CUPEDIT', 'CUPDEL', 'CUPOPEN', 'CUPDRAW', 'CUPSTART', 'CUPADD', 'CUPREMOVE',
                     'CUPSWAP', 'CUPRESULT')
        if cmd == 'CUPS':
            c.caps.add('cups')
            self.send_cups(c)
            return True
        if cmd in host_only and not c.admin:
            c.send('ERR', 'NOT_ADMIN', 'Host mode is required to manage tournaments. Use HOST|password first.')
            return True
        if cmd == 'CUPNEW' or cmd == 'CUPEDIT':
            k = 2 if cmd == 'CUPEDIT' else 1
            t = None
            if cmd == 'CUPEDIT':
                t = self.find_cup(arg(1))
                if not t:
                    c.send('ERR', 'CUP_NOT_FOUND', 'Tournament not found. It may have been deleted.')
                    return True
            elif len(self.cups) >= MAX_CUPS:
                c.send('ERR', 'CUP_FULL', 'There can be up to %d tournaments. Delete a tournament first.' % MAX_CUPS)
                return True
            try:
                most = int(arg(k) or 0)
            except ValueError:
                most = 0
            if most and not 2 <= most <= MAX_CUP_PLAYERS:
                c.send('ERR', 'CUP_BAD', 'The maximum number of players must be from 2 to %d.' % MAX_CUP_PLAYERS)
                return True
            name = clean(arg(k + 3), 30).replace('~', '-')
            if not name:
                c.send('ERR', 'CUP_BAD', 'Tournament name is required.')
                return True
            if t and t['state'] == 'signup' and most and most < len(t['players']):
                c.send('ERR', 'CUP_BAD', '%d players have already signed up. The maximum cannot be lower than that.' % len(t['players']))
                return True
            fields = {'max': most, 'time': clean(arg(k + 1), 12), 'game': clean(arg(k + 2), 40), 'name': name,
                      'notes': clean(arg(k + 4, rest=True), MAX_NOTE).replace('|', '/').strip('~ ')}
            if t:
                t.update(fields)
                self.cup_changed(t)
                c.send('CUPSAVED', t['id'])
                log('%s changed the tournament %s' % (c.nick, name))
                return True
            t = {'id': max([x['id'] for x in self.cups] + [0]) + 1, 'state': 'signup', 'players': [], 'slots': [],
                 'results': {}, 'reports': {}, 'champion': '', 'by': c.nick}
            t.update(fields)
            self.cups.append(t)
            self.cup_changed(t)
            c.send('CUPSAVED', t['id'])
            game = ' (%s)' % t['game'] if t['game'] else ''
            self.note('trophy', 'Tournament sign-ups open: %s%s. See Tournaments.' % (name, game))
            log('%s created the tournament %s' % (c.nick, name))
            return True
        t = self.find_cup(arg(1))
        if not t:
            c.send('ERR', 'CUP_NOT_FOUND', 'Tournament not found. It may have been deleted.')
            return True
        if cmd == 'CUPDEL':
            self.cups.remove(t)
            self._save_cups()
            self.broadcast('CUPDEL', t['id'])
            if t['state'] != 'done':
                self.note('trophy', '%s has been cancelled.' % t['name'])
            log('%s deleted the tournament %s' % (c.nick, t['name']))
            return True
        if cmd in ('CUPJOIN', 'CUPLEAVE', 'CUPADD', 'CUPREMOVE'):
            nick = c.nick if cmd in ('CUPJOIN', 'CUPLEAVE') else clean(arg(2), 12)
            if t['state'] != 'signup':
                if cmd in ('CUPJOIN', 'CUPLEAVE'):
                    c.send('ERR', 'CUP_CLOSED', 'Sign-ups for %s are closed. Ask the host.' % t['name'])
                else:
                    c.send('ERR', 'CUP_CLOSED', 'The bracket has been drawn. Click Reopen Sign-ups to change the players.')
                return True
            if not NICK_RE.match(nick):
                c.send('ERR', 'CUP_BAD', 'Invalid nickname.')
                return True
            there = next((p for p in t['players'] if p.lower() == nick.lower()), None)
            if cmd in ('CUPJOIN', 'CUPADD'):
                if there:
                    return True
                if t['max'] and len(t['players']) >= t['max']:
                    c.send('ERR', 'CUP_FULL', '%s is full (%d players).' % (t['name'], t['max']))
                    return True
                if len(t['players']) >= MAX_CUP_PLAYERS:
                    c.send('ERR', 'CUP_FULL', 'A tournament can have up to %d players.' % MAX_CUP_PLAYERS)
                    return True
                t['players'].append(nick)
            else:
                if not there:
                    return True
                t['players'].remove(there)
            self.cup_changed(t)
            return True
        if cmd == 'CUPOPEN':
            if t['state'] in ('running', 'done'):
                c.send('ERR', 'CUP_STARTED', '%s has started. Sign-ups cannot be reopened.' % t['name'])
                return True
            t['state'], t['slots'], t['results'], t['reports'] = 'signup', [], {}, {}
            self.cup_changed(t)
            self.note('trophy', 'Tournament sign-ups open again: %s. See Tournaments.' % t['name'])
            return True
        if cmd == 'CUPDRAW':
            if t['state'] not in ('signup', 'drawn'):
                c.send('ERR', 'CUP_STARTED', '%s has started. The bracket cannot be drawn again.' % t['name'])
                return True
            if len(t['players']) < 2:
                c.send('ERR', 'CUP_FEW', 'At least 2 players are required to draw the bracket.')
                return True
            first = t['state'] == 'signup'
            self.cup_draw(t)
            self.cup_changed(t)
            if first:
                self.note('trophy', '%s: sign-ups closed and the bracket has been drawn. See Tournaments.' % t['name'])
            log('%s drew the bracket for %s' % (c.nick, t['name']))
            return True
        if cmd == 'CUPSWAP':
            if t['state'] != 'drawn':
                c.send('ERR', 'CUP_STARTED', 'Players can only be swapped before the tournament starts.')
                return True
            low = [p.lower() if p else None for p in t['slots']]
            a, b = arg(2).lower(), arg(3).lower()
            if a not in low or b not in low or a == b:
                c.send('ERR', 'CUP_BAD', 'Select two different players in the bracket.')
                return True
            ia, ib = low.index(a), low.index(b)
            t['slots'][ia], t['slots'][ib] = t['slots'][ib], t['slots'][ia]
            self.cup_changed(t)
            return True
        if cmd == 'CUPSTART':
            if t['state'] != 'drawn':
                c.send('ERR', 'CUP_BAD', 'Draw the bracket first.')
                return True
            t['state'] = 'running'
            self.cup_changed(t, before=set())
            n = sum(1 for p in t['slots'] if p)
            self.note('trophy', '%s has started (%d players). See Tournaments for your first match.' % (t['name'], n))
            log('%s started the tournament %s' % (c.nick, t['name']))
            return True
        if cmd in ('CUPREPORT', 'CUPRESULT'):
            if t['state'] not in ('running', 'done'):
                c.send('ERR', 'CUP_BAD', '%s has not started.' % t['name'])
                return True
            m, rounds = self.cup_match(t, arg(2), arg(3))
            if not m or not m['a'] or not m['b'] or m['auto']:
                c.send('ERR', 'CUP_BAD', 'This match cannot have a result yet.')
                return True
            side = arg(4).lower()
            score = clean(arg(5), 20).replace('|', '/')
            before = self.cup_ready_now(t)
            nxt = rounds[m['r'] + 1][m['i'] // 2] if m['r'] + 1 < len(rounds) else None
            if cmd == 'CUPREPORT':
                if c.nick.lower() not in (m['a'].lower(), m['b'].lower()):
                    c.send('ERR', 'CUP_BAD', 'Only the two players can report the result of a match.')
                    return True
                if m['win']:
                    c.send('ERR', 'CUP_DONE', 'The result of this match has already been confirmed. Ask the host to change it.')
                    return True
                if side not in ('a', 'b'):
                    c.send('ERR', 'CUP_BAD', 'Select the winner.')
                    return True
                t['reports'][m['key']] = {'by': c.nick, 'win': side, 'score': score, 'a': m['a'], 'b': m['b']}
                self.cup_changed(t)
                for h in self.signed_in_clients():
                    if h.admin and 'cups' in h.caps:
                        h.send('CUPREPORTED', t['id'], m['r'], m['i'], c.nick)
                log('%s reported a result in %s: %s' % (c.nick, t['name'], m[side]))
                return True
            # the host: confirm, enter, change or clear a result
            if side not in ('a', 'b', ''):
                c.send('ERR', 'CUP_BAD', 'Select the winner.')
                return True
            changing_winner = (side != m['win'])
            if changing_winner and nxt and nxt['win'] and not nxt['auto']:
                c.send('ERR', 'CUP_NEXT', 'The next match already has a result. Clear that result first.')
                return True
            if side:
                t['results'][m['key']] = {'win': side, 'score': score, 'a': m['a'], 'b': m['b']}
            else:
                t['results'].pop(m['key'], None)
            t['reports'].pop(m['key'], None)
            self.cup_changed(t, before=before)
            if side and changing_winner:
                loser = m['b'] if side == 'a' else m['a']
                rname = self.round_name(m['r'], len(rounds))
                self.note('trophy', '%s, %s: %s beat %s%s.' % (t['name'], rname, m[side], loser, (' ' + score) if score else ''))
                if t['state'] == 'done':
                    self.note('trophy', '%s is the winner of %s.' % (t['champion'], t['name']))
                    self.broadcast('CUPWIN', t['id'], t['champion'])
                    log('%s won the tournament %s' % (t['champion'], t['name']))
            return True
        return False

    def note(self, icon, text):
        """A line from the LAN Center itself in everyone's chat."""
        self.broadcast('NOTE', icon, text)
        self.chat_log('[%s]' % text)

    def start_vote(self, c, minutes, question, options):
        if self.vote and self.vote.get('task'):
            self.vote['task'].cancel()
        loop = asyncio.get_running_loop()
        self.vote = {'question': question, 'options': options, 'votes': {}, 'state': 'open',
                     'deadline': loop.time() + minutes * 60, 'by': c.nick, 'winner': 0}
        self.vote['task'] = loop.create_task(self._vote_timer(minutes * 60))
        self.vote_changed()
        used = {o.lower() for o in options}
        if any(s['name'].lower() in used for s in self.suggest):
            self.suggest = [s for s in self.suggest if s['name'].lower() not in used]
            self.suggestions_changed()
        self.note('star', 'Vote started: %s Voting closes in %d minute%s. See What\'s On.'
                  % (question, minutes, '' if minutes == 1 else 's'))
        log('%s started a vote: %s %s (%d min)' % (c.nick, question, options, minutes))

    async def _vote_timer(self, seconds):
        await asyncio.sleep(seconds)
        self.vote['task'] = None
        self.close_vote()

    def close_vote(self):
        v = self.vote
        if not v or v['state'] != 'open':
            return
        if v.get('task'):
            v['task'].cancel()
            v['task'] = None
        v['state'] = 'closed'
        counts = self.vote_counts()
        total = sum(counts)
        best = max(counts) if counts else 0
        leaders = [v['options'][i] for i, n in enumerate(counts) if n == best]
        if total == 0:
            v['winner'] = 0
            text = 'Voting closed. No votes were cast.'
        elif len(leaders) > 1:
            v['winner'] = 0
            text = 'Voting closed. Tie between %s.' % ' and '.join(leaders)
        else:
            v['winner'] = counts.index(best) + 1
            text = 'Voting closed. %s won with %d of %d vote%s.' % (leaders[0], best, total, '' if total == 1 else 's')
        self.vote_changed()
        self.note('star', text)
        log('Vote closed: %s' % text)

    # ----- seat maps, one per venue
    def _load_maps(self):
        try:
            d = json.loads(MAPS_FILE.read_text(encoding='utf-8'))
            return d.get('maps', []), int(d.get('active', 0))
        except FileNotFoundError:
            maps, active = [json.loads(json.dumps(DEFAULT_MAP))], 1
            try:
                MAPS_FILE.write_text(json.dumps({'active': active, 'maps': maps}, indent=1), encoding='utf-8')
            except OSError:
                pass
            return maps, active
        except Exception as e:
            log('Could not read %s (%s). Using the default seat map.' % (MAPS_FILE, e))
            return [json.loads(json.dumps(DEFAULT_MAP))], 1

    def _save_maps(self):
        tmp = MAPS_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps({'active': self.map_active, 'maps': self.maps}, indent=1), encoding='utf-8')
        os.replace(tmp, MAPS_FILE)

    @staticmethod
    def map_wire(m):
        """name, stage, columns, tables -- tables as  Table A,A,4,2~Table B,B,4,2"""
        tables = '~'.join('%s,%s,%d,%d' % (t['name'], t['prefix'], t['per_side'], t['sides']) for t in m['tables'])
        return m['name'], m['stage'], m['columns'], tables

    def find_map(self, map_id):
        return next((m for m in self.maps if str(m['id']) == str(map_id)), None)

    def send_map(self, c):
        m = self.find_map(self.map_active)
        if m:
            c.send('MAP', m['id'], *self.map_wire(m))
        else:
            c.send('MAP', 0, '', '', 1, '')

    def map_changed(self):
        for c in self.signed_in_clients():
            self.send_map(c)

    def send_map_list(self, c):
        for m in self.maps:
            c.send('MAPL', m['id'], 1 if m['id'] == self.map_active else 0, m['name'])
        c.send('MAPLEND', len(self.maps))

    @staticmethod
    def parse_map(name, stage, columns, raw):
        """Checks a seat map from the editor. Returns (map, None) or (None, what's wrong)."""
        tidy = lambda x, n: clean(x.replace('~', '-').replace(',', ' '), n)
        try:
            columns = max(1, min(4, int(columns)))
        except ValueError:
            columns = 1
        tables, letters = [], set()
        for part in raw.split('~'):
            bits = part.split(',')
            if len(bits) < 4:
                continue
            prefix = bits[1].strip().upper()
            if not PREFIX_RE.match(prefix):
                return None, 'Each table must have a seat letter (for example, A or BB).'
            if prefix in letters:
                return None, 'Seat letter %s is used by more than one table. Each table must have its own seat letter.' % prefix
            letters.add(prefix)
            try:
                per = int(bits[2])
            except ValueError:
                per = 0
            if not 1 <= per <= 12:
                return None, 'Tables can have 1 to 12 seats per side.'
            tables.append({'name': tidy(bits[0], 20) or 'Table ' + prefix, 'prefix': prefix, 'per_side': per,
                           'sides': 2 if bits[3].strip() == '2' else 1})
        if not tables:
            return None, 'A seat map must have at least one table.'
        if len(tables) > 26:
            return None, 'A seat map can have a maximum of 26 tables.'
        return {'name': tidy(name, 30) or 'New Venue', 'stage': tidy(stage, 40), 'columns': columns, 'tables': tables}, None

    # ----- the Game Library: the pak server's list of games
    def send_paks(self, c):
        """PAKS|how many paks|pak server address|note, a PAK line per pak (with a GAME line
        for each game in a pack after it), then PAKEND"""
        n = sum(1 for l in self.paks if l.startswith('PAK|'))
        c.send('PAKS', n, self.paks_base, self.paks_note)
        for line in self.paks:
            c.send(line)
        c.send('PAKEND', n)

    def paks_changed(self):
        for c in self.signed_in_clients():
            if 'paks' in c.caps:
                self.send_paks(c)

    async def pak_watch(self):
        """Every few seconds, fetch catalog.txt from the pak server. Clients hear only about changes."""
        url = self.cfg.get('paks_url', '').strip()
        if not url:
            self.paks_note = 'No Cartridge Server is configured (paks_url in palms.ini).'
            log('paks_url is not set in palms.ini. The Game Library is disabled.')
            return
        if not url.endswith('/'):
            url += '/'
        trouble = None
        while True:
            try:
                text = await asyncio.to_thread(fetch_text, url + 'catalog.txt')
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                if not lines or not lines[0].startswith('PAKS|'):
                    raise ValueError('invalid catalog format')
                head = lines[0].split('|')
                base = head[4].strip() if len(head) > 4 and head[4].startswith('http') else url
                if not base.endswith('/'):
                    base += '/'
                paks = [l for l in lines[1:] if l.startswith(('PAK|', 'GAME|'))]
                note = '' if paks else 'No games are available on the Cartridge Server.'
                if trouble:
                    log('Cartridge Server at %s is responding again.' % url)
                    trouble = None
                if paks != self.paks or base != self.paks_base or note != self.paks_note:
                    self.paks, self.paks_base, self.paks_note = paks, base, note
                    n = sum(1 for l in paks if l.startswith('PAK|'))
                    log('Game Library: %d cartridge%s on the Cartridge Server.' % (n, '' if n == 1 else 's'))
                    self.paks_changed()
            except Exception as e:
                why = str(e) or e.__class__.__name__
                if trouble != why:
                    log('Unable to retrieve %scatalog.txt (%s). Retrying every %d seconds.'
                        % (url, why, PAK_SECONDS))
                    trouble = why
                note = 'Unable to connect to the Cartridge Server at %s.' % url
                if self.paks_note != note:
                    # keep the games we already know about: they may still download
                    self.paks_note = note
                    self.paks_changed()
            await asyncio.sleep(PAK_SECONDS)

    # ----- who's here
    @staticmethod
    def hm_at(t):
        return datetime.datetime.fromtimestamp(t).strftime('%H:%M') if t else ''

    def forget_old(self):
        cutoff = time.time() - PARTY_HOURS * 3600
        self.arrived = {k: t for k, t in self.arrived.items() if t >= cutoff}
        self.gone = {k: g for k, g in self.gone.items() if g['t'] >= cutoff}

    def send_who(self, c):
        self.forget_old()
        people = sorted(self.signed_in_clients(), key=lambda x: x.nick.lower())
        for p in people:
            c.send('WHO', p.nick, p.seat or '-', p.status, 1 if p.admin else 0, self.hm_at(self.arrived.get(p.nick.lower())),
                   p.addr, '' if p.ms is None else p.ms, p.specs[7])
        gone = sorted(self.gone.values(), key=lambda g: -g['t'])[:MAX_GONE]
        for g in gone:
            c.send('GONE', g['nick'], g['seat'] or '-', g['how'], g['at'])
        c.send('WHOEND', len(people))

    def send_whois(self, c, nick):
        f = lambda x: str(x or '').replace('|', '/')
        p = self.find_nick(nick)
        if p:
            c.send('WHOIS', p.nick, p.seat or '-', p.status, 1 if p.admin else 0,
                   self.hm_at(self.arrived.get(p.nick.lower())), '', f(p.pcname), p.addr, f(p.ver),
                   *[f(x) for x in p.specs], '' if p.ms is None else p.ms)
            return
        g = self.gone.get(nick.lower())
        if g:
            specs = (list(g['specs']) + [''] * 8)[:8]
            c.send('WHOIS', g['nick'], g['seat'] or '-', g['how'], 0, g['arrived'], g['at'],
                   f(g['pc']), g['ip'], f(g['ver']), *[f(x) for x in specs], '')
            return
        c.send('WHOIS', clean(nick, 12).replace('|', ''), '-', 'unknown', 0, *([''] * 14))

    def left(self, c, how):
        """Remember who left, so Who's Here can show them as gone home, on a break, and so on."""
        key = c.nick.lower()
        self.gone[key] = {'nick': c.nick, 'seat': c.seat, 'how': how, 'at': now_hm(), 't': time.time(),
                          'arrived': self.hm_at(self.arrived.get(key)), 'pc': c.pcname, 'ip': c.addr,
                          'ver': c.ver, 'specs': list(c.specs)}
        if how == 'home':
            self.arrived.pop(key, None)
            changed = False
            for e in self.schedule:
                if any(n.lower() == key for n in e.get('in', [])):
                    e['in'] = [n for n in e['in'] if n.lower() != key]
                    changed = True
            if changed:
                self.schedule_changed()
            for t in self.cups:
                if t['state'] == 'signup' and any(n.lower() == key for n in t['players']):
                    t['players'] = [n for n in t['players'] if n.lower() != key]
                    self.cup_changed(t)

    def clear_held(self, seat, nick):
        """Someone sat in a seat that was being kept for a person who left: it's theirs now."""
        for g in self.gone.values():
            if seat and g['seat'] == seat and g['nick'].lower() != nick.lower():
                g['seat'] = None

    def change_seat(self, c, seat):
        if seat in ('', '-'):
            seat = None
        elif not SEAT_RE.match(seat):
            c.send('ERR', 'BAD_SEAT', 'Invalid seat number. Seat numbers look like A01 or B12.')
            return
        if seat and seat != c.seat:
            sitter = self.find_seat(seat)
            if sitter and sitter is not c:
                c.send('ERR', 'SEAT_TAKEN', 'Seat %s is taken by %s.' % (seat, sitter.nick))
                return
        old, c.seat = c.seat, seat
        self.clear_held(seat, c.nick)
        if c.pcid and c.pcid in self.pcs:
            self.pcs[c.pcid]['seat'] = seat or ''
            self._save_pcs()
        self.broadcast('SEATCHG', c.nick, seat or '-')
        if old != seat:
            log('%s moved from seat %s to %s.' % (c.nick, old or '-', seat or '-'))

    # ----- helpers
    def signed_in_clients(self):
        return [c for c in self.clients if c.signed_in]

    def find_nick(self, nick):
        for c in self.signed_in_clients():
            if c.nick.lower() == nick.lower():
                return c
        return None

    def find_seat(self, seat):
        for c in self.signed_in_clients():
            if c.seat == seat:
                return c
        return None

    def broadcast(self, *fields, skip=None):
        for c in self.signed_in_clients():
            if c is not skip:
                c.send(*fields)

    def chat_log(self, line):
        path = DATA_DIR / ('chat-%s.log' % datetime.date.today().isoformat())
        with open(path, 'a', encoding='utf-8') as f:
            f.write('%s %s\n' % (datetime.datetime.now().strftime('%H:%M:%S'), line))

    # ----- connections
    async def handle(self, reader, writer):
        c = Client(self, reader, writer)
        self.clients.add(c)
        log('Connected: %s' % c.addr)
        c.send('INFO', 'Welcome to %s. Type HELP for a list of commands.' % self.name)
        reason = 'dropped'
        try:
            while not c.closed:
                timeout = REAL_CLIENT_TIMEOUT if c.pcid else None
                try:
                    raw = await asyncio.wait_for(reader.readline(), timeout)
                except asyncio.TimeoutError:
                    log('%s timed out after %d seconds. Disconnecting.' % (c.label(), REAL_CLIENT_TIMEOUT))
                    break
                except (asyncio.LimitOverrunError, ValueError):
                    c.send('ERR', 'TOO_LONG', 'Line too long.')
                    break
                if not raw:
                    break
                line = raw[:MAX_LINE].decode('cp1252', 'replace').rstrip('\r\n')
                line = line.replace('\x08', '')
                if not line.strip():
                    continue
                result = self.command(c, line)
                if result:
                    reason = result
                    break
                await writer.drain()
        except (ConnectionError, OSError):
            pass
        finally:
            self.clients.discard(c)
            c.closed = True
            if c.signed_in:
                self.left(c, reason)
                if reason == 'break':
                    self.broadcast('STATUS', c.nick, 'away')
                else:
                    self.broadcast('LEFT', c.nick, reason)
                log('%s left (%s).' % (c.nick, reason))
            else:
                log('Disconnected: %s' % c.addr)
            try:
                writer.close()
                await writer.wait_closed()
            except Exception:
                pass

    # ----- the protocol. Returns a reason string when the connection should close.
    def command(self, c, line):
        parts = line.split('|')
        cmd = parts[0].strip().upper()

        def arg(i, rest=False):
            if rest:
                return '|'.join(parts[i:]) if len(parts) > i else ''
            return parts[i].strip() if len(parts) > i else ''

        if cmd == 'HELP':
            for h in HELP_TEXT:
                c.send('INFO', h)
            return None
        if cmd == 'PING':
            c.send('PONG', datetime.datetime.now().strftime('%H:%M:%S'))
            return None
        if cmd == 'QUIT':
            c.send('BYE', 'Goodbye.')
            if arg(1).lower() == 'update':
                # back in a few seconds with the new version: see login()
                if c.signed_in:
                    self.updating[c.nick.lower()] = time.time()
                return 'update'
            return 'quit'
        if cmd == 'HELLO':
            c.pcid = clean(arg(1), 40) or None
            c.ver = clean(arg(2), 12)
            c.pcname = clean(arg(3), 20)
            c.send('WELCOME', self.name, VERSION)
            known = self.pcs.get(c.pcid) if c.pcid else None
            if known:
                c.send('KNOWN', known.get('nick', ''), known.get('seat', ''))
            else:
                c.send('NEWPC')
            self.offer_update(c)
            return None
        if cmd == 'LOGIN':
            return self.login(c, arg(1), arg(2).upper())
        if cmd == 'CAPS':
            c.caps = {x.strip().lower() for x in parts[1:] if x.strip()}
            return None
        if cmd == 'SPONG':
            try:
                ms = int(asyncio.get_running_loop().time() * 1000) - int(arg(1))
            except ValueError:
                return None
            if 0 <= ms < 60000:
                c.ms = ms
            return None

        if not c.signed_in:
            c.send('ERR', 'NOT_SIGNED_IN', 'Not signed in. Use LOGIN|nickname|seat (for example, LOGIN|Player1|A01).')
            return None

        if cmd == 'CHAT':
            text = clean(arg(1, rest=True))
            if not text:
                return None
            if c.status == 'away':
                c.status = 'here'
                self.broadcast('STATUS', c.nick, 'here')
            t = now_hm()
            self.history.append((t, c.nick, text))
            self.broadcast('CHAT', t, c.nick, text)
            self.chat_log('<%s> %s' % (c.nick, text))
            return None
        if cmd == 'WHO':
            self.send_who(c)
            return None
        if cmd == 'WHOIS':
            self.send_whois(c, arg(1))
            return None
        if cmd == 'SEAT':
            self.change_seat(c, arg(1).upper())
            return None
        if cmd == 'SPECS':
            c.specs = [clean(arg(i), 60) for i in range(1, 9)]
            return None
        if cmd == 'MAP':
            self.send_map(c)
            return None
        if cmd == 'PAKS':
            self.send_paks(c)
            return None
        if cmd in ('MAPLIST', 'MAPGET', 'MAPSAVE', 'MAPDEL', 'MAPUSE') and not c.admin:
            c.send('ERR', 'NOT_ADMIN', 'Host mode is required to change the seat map. Use HOST|password first.')
            return None
        if cmd == 'MAPLIST':
            self.send_map_list(c)
            return None
        if cmd == 'MAPGET':
            m = self.find_map(arg(1))
            if not m:
                c.send('ERR', 'NOT_FOUND', 'Venue not found. It may have been deleted.')
                return None
            c.send('MAPDEF', m['id'], *self.map_wire(m))
            return None
        if cmd == 'MAPSAVE':
            new, problem = self.parse_map(arg(2), arg(3), arg(4), arg(5))
            if problem:
                c.send('ERR', 'MAP_BAD', problem)
                return None
            m = self.find_map(arg(1))
            if not m:
                m = {'id': max([x['id'] for x in self.maps] + [0]) + 1}
                self.maps.append(m)
            m.update(new)
            self._save_maps()
            c.send('MAPSAVED', m['id'], m['name'])
            self.send_map_list(c)
            if m['id'] == self.map_active:
                self.map_changed()
            log('%s saved the seat map %s' % (c.nick, m['name']))
            return None
        if cmd == 'MAPDEL':
            m = self.find_map(arg(1))
            if m and m['id'] == self.map_active:
                c.send('ERR', 'MAP_INUSE', 'This seat map is in use and cannot be deleted. Use another venue first.')
                return None
            self.maps = [x for x in self.maps if x is not m]
            self._save_maps()
            self.send_map_list(c)
            return None
        if cmd == 'MAPUSE':
            m = self.find_map(arg(1))
            if not m:
                c.send('ERR', 'NOT_FOUND', 'Venue not found. It may have been deleted.')
                return None
            self.map_active = m['id']
            self._save_maps()
            self.map_changed()
            self.send_map_list(c)
            self.note('plan', '%s changed the seat map to %s. See Who\'s Here.' % (c.nick, m['name']))
            log('%s is using the seat map %s' % (c.nick, m['name']))
            return None
        if cmd in ('AWAY', 'BUSY', 'BACK'):
            c.status = {'AWAY': 'away', 'BUSY': 'busy', 'BACK': 'here'}[cmd]
            self.broadcast('STATUS', c.nick, c.status)
            return None
        if cmd in ('ADMIN', 'HOST'):
            if hmac.compare_digest(arg(1).encode(), self.admin_password.encode()):
                c.admin = True
                c.send('ADMINOK', 'Host mode enabled.')
                log('%s enabled host mode.' % c.nick)
            else:
                c.send('ERR', 'BAD_PASSWORD', 'Incorrect host password.')
                log('Incorrect host password from %s (%s).' % (c.nick, c.addr))
            return None
        if cmd == 'PA':
            if not c.admin:
                c.send('ERR', 'NOT_ADMIN', 'Host mode is required to send announcements. Use HOST|password first.')
                return None
            chime = arg(1).lower()
            text = clean(arg(2, rest=True))
            if chime not in CHIMES:
                c.send('ERR', 'BAD_CHIME', 'Chime must be dingdong, triple, horn or none.')
                return None
            if not text:
                c.send('ERR', 'EMPTY', 'Message text is required (for example, PA|dingdong|Pizza is here).')
                return None
            self.broadcast('PA', chime, c.nick, now_hm(), text)
            self.chat_log('[PA from %s] %s' % (c.nick, text))
            log('Announcement from %s: %s' % (c.nick, text))
            return None
        if cmd == 'SCHED':
            self.send_schedule(c)
            return None
        if cmd == 'VOTEGET':
            self.send_vote(c)
            return None
        if cmd == 'VOTE':
            v = self.vote
            if not v or v['state'] != 'open':
                c.send('ERR', 'NO_VOTE', 'No vote in progress.')
                return None
            try:
                n = int(arg(1))
            except ValueError:
                n = 0
            if not 1 <= n <= len(v['options']):
                c.send('ERR', 'BAD_VOTE', 'Choice must be from 1 to %d.' % len(v['options']))
                return None
            v['votes'][c.nick.lower()] = n - 1
            self.vote_changed()
            return None
        if cmd in ('SCHEDIN', 'SCHEDOUT'):
            e = next((e for e in self.schedule if str(e['id']) == arg(1)), None)
            if not e:
                c.send('ERR', 'NOT_FOUND', 'Event %s not found. Type SCHED to list events.' % clean(arg(1), 10))
                return None
            ins = [n for n in e.get('in', []) if n.lower() != c.nick.lower()]
            if cmd == 'SCHEDIN':
                ins.append(c.nick)
            if ins != e.get('in', []):
                e['in'] = ins
                self.schedule_changed()
            return None
        if cmd == 'SUGGEST':
            name = clean(arg(1, rest=True), 40).replace('|', '/').replace(',', ' ').strip()
            if not name:
                c.send('ERR', 'EMPTY', 'Game name is required (for example, SUGGEST|Quake III Arena).')
                return None
            s = self.find_suggestion(name)
            if s and c.nick in s['by']:
                return None
            backed = sum(1 for x in self.suggest if c.nick in x['by'])
            if backed >= MAX_BACKED:
                c.send('ERR', 'TOO_MANY', 'You can suggest up to %d games. Withdraw one first.' % MAX_BACKED)
                return None
            if s:
                s['by'].append(c.nick)
            else:
                if len(self.suggest) >= MAX_SUGGESTIONS:
                    c.send('ERR', 'FULL', 'The suggestion list is full (%d games). Back a game already on the list.' % MAX_SUGGESTIONS)
                    return None
                self.suggest.append({'name': name, 'by': [c.nick]})
                self.note('star', '%s suggested %s for the next vote.' % (c.nick, name))
                log('%s suggested %s' % (c.nick, name))
            self.suggestions_changed()
            return None
        if cmd == 'UNSUGGEST':
            s = self.find_suggestion(clean(arg(1, rest=True), 40))
            if s and c.nick in s['by']:
                s['by'].remove(c.nick)
                self.suggestions_changed()
            return None
        if cmd == 'FOOD':
            f = self.food
            if f['state'] != 'open':
                c.send('ERR', 'FOOD_CLOSED', 'No food order is open.')
                return None
            text = clean(arg(1, rest=True), 80).replace('|', '/')
            orders = [o for o in f['orders'] if o['nick'].lower() != c.nick.lower()]
            if text:
                mine = next((o for o in f['orders'] if o['nick'].lower() == c.nick.lower()), None)
                o = {'nick': c.nick, 'time': now_hm(), 'text': text}
                if mine:
                    orders.insert(f['orders'].index(mine), o)
                else:
                    orders.append(o)
                log('Food order from %s: %s' % (c.nick, text))
            f['orders'] = orders
            self.food_changed()
            return None
        if cmd.startswith('TPL') and cmd in ('TPLLIST', 'TPLGET', 'TPLSAVE', 'TPLDEL', 'TPLAPPLY') and not c.admin:
            c.send('ERR', 'NOT_ADMIN', 'Host mode is required to use schedule templates. Use HOST|password first.')
            return None
        if cmd == 'TPLLIST':
            self.send_template_list(c)
            return None
        if cmd == 'TPLGET':
            t = next((t for t in self.templates if str(t['id']) == arg(1)), None)
            if not t:
                c.send('ERR', 'NOT_FOUND', 'Template not found. It may have been deleted.')
                return None
            c.send('TPLTEXT', t['id'], t['name'], '~'.join(t['lines']))
            return None
        if cmd == 'TPLSAVE':
            name = clean(arg(2).replace('~', '-'), 40) or 'Untitled'
            lines = self.template_lines(arg(3, rest=True))
            t = next((t for t in self.templates if str(t['id']) == arg(1)), None)
            if not t:
                t = {'id': max([x['id'] for x in self.templates] + [0]) + 1}
                self.templates.append(t)
            t['name'], t['lines'] = name, lines
            self._save_templates()
            c.send('TPLSAVED', t['id'], name)
            self.send_template_list(c)
            log('%s saved the template %s' % (c.nick, name))
            return None
        if cmd == 'TPLDEL':
            self.templates = [t for t in self.templates if str(t['id']) != arg(1)]
            self._save_templates()
            self.send_template_list(c)
            return None
        if cmd == 'TPLAPPLY':
            self.apply_template(c, 'add' if arg(1).lower() == 'add' else 'replace', self.template_lines(arg(2, rest=True)))
            return None
        if cmd in ('SCHEDADD', 'SCHEDDEL', 'SCHEDNOTE', 'VOTESTART', 'VOTESTOP', 'SUGGDEL',
                   'FOODOPEN', 'FOODCLOSE', 'FOODCLEAR') and not c.admin:
            c.send('ERR', 'NOT_ADMIN', 'Host mode is required to change What\'s On. Use HOST|password first.')
            return None
        if cmd == 'SCHEDADD':
            m = TIME_RE.match(arg(1))
            name = clean(arg(3, rest=True), 40)
            if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
                c.send('ERR', 'BAD_TIME', 'Invalid time. Use 24-hour format (for example, 19:30).')
                return None
            if not name:
                c.send('ERR', 'EMPTY', 'Event description is required (for example, SCHEDADD|19:30|trophy|Quake 2 final).')
                return None
            icon = arg(2).lower() if arg(2).lower() in EVENT_ICONS else 'star'
            new_id = max([e['id'] for e in self.schedule] + [0]) + 1
            self.schedule.append({'id': new_id, 'time': '%02d:%s' % (int(m.group(1)), m.group(2)),
                                  'icon': icon, 'name': name})
            self.schedule_changed()
            c.send('SCHEDADDED', new_id)
            log('%s added to the schedule: %s %s' % (c.nick, arg(1), name))
            return None
        if cmd == 'SCHEDDEL':
            before = len(self.schedule)
            self.schedule = [e for e in self.schedule if str(e['id']) != arg(1)]
            if len(self.schedule) == before:
                c.send('ERR', 'NOT_FOUND', 'Event %s not found. Type SCHED to list events.' % clean(arg(1), 10))
                return None
            self.schedule_changed()
            return None
        if cmd == 'SCHEDNOTE':
            e = next((e for e in self.schedule if str(e['id']) == arg(1)), None)
            if not e:
                c.send('ERR', 'NOT_FOUND', 'Event %s not found. Type SCHED to list events.' % clean(arg(1), 10))
                return None
            note = clean(arg(2, rest=True), MAX_NOTE).replace('|', '/').strip('~ ')
            if note:
                e['note'] = note
            else:
                e.pop('note', None)
            self.schedule_changed()
            return None
        if cmd == 'SUGGDEL':
            s = self.find_suggestion(clean(arg(1, rest=True), 40))
            if not s:
                c.send('ERR', 'NOT_FOUND', 'That game is not on the suggestion list.')
                return None
            self.suggest.remove(s)
            self.suggestions_changed()
            return None
        if cmd == 'FOODOPEN':
            f = self.food
            f['title'] = clean(arg(1, rest=True), 60).replace('|', '/') or 'Food order'
            was = f['state']
            f['state'] = 'open'
            self.food_changed()
            if was != 'open':
                self.note('pizza', 'Food order open: %s. Click Food Order on the What\'s On page to order.' % f['title'])
            log('%s opened the food order: %s' % (c.nick, f['title']))
            return None
        if cmd == 'FOODCLOSE':
            f = self.food
            if f['state'] != 'open':
                c.send('ERR', 'FOOD_CLOSED', 'No food order is open.')
                return None
            f['state'] = 'closed'
            self.food_changed()
            n = len(f['orders'])
            self.note('pizza', 'Food order closed: %d order%s.' % (n, '' if n == 1 else 's'))
            log('%s closed the food order (%d orders)' % (c.nick, n))
            return None
        if cmd == 'FOODCLEAR':
            self.food['orders'] = []
            self.food_changed()
            log('%s cleared the food orders' % c.nick)
            return None
        if cmd == 'VOTESTART':
            try:
                minutes = int(arg(1))
            except ValueError:
                minutes = 0
            question = clean(arg(2), 80) or 'Next game?'
            options = [clean(o, 40) for o in parts[3:]]
            options = [o for o in options if o][:MAX_VOTE_OPTIONS]
            if not 1 <= minutes <= 60:
                c.send('ERR', 'BAD_MINUTES', 'Vote duration must be from 1 to 60 minutes.')
                return None
            if len(options) < 2:
                c.send('ERR', 'FEW_CHOICES', 'A vote must have at least 2 choices.')
                return None
            self.start_vote(c, minutes, question, options)
            return None
        if cmd == 'VOTESTOP':
            if not self.vote or self.vote['state'] != 'open':
                c.send('ERR', 'NO_VOTE', 'No vote in progress.')
                return None
            self.close_vote()
            return None
        if cmd.startswith('CUP') and self.cup_command(c, cmd, parts, arg):
            return None
        if cmd == 'LOGOUT':
            how = arg(1).lower() or 'home'
            if how not in ('break', 'home'):
                how = 'home'
            if how == 'home' and c.pcid and c.pcid in self.pcs:
                self.pcs[c.pcid]['seat'] = ''
                self._save_pcs()
            c.send('BYE', 'Signed out. Your seat has been kept.' if how == 'break' else 'Signed out.')
            return how

        c.send('ERR', 'UNKNOWN', 'Unknown command "%s". Type HELP for a list of commands.' % clean(cmd, 20))
        return None

    def login(self, c, nick, seat):
        if c.signed_in:
            c.send('ERR', 'ALREADY', 'You are already signed in as %s.' % c.nick)
            return None
        if not NICK_RE.match(nick):
            c.send('ERR', 'BAD_NICK', 'Nicknames must be 2 to 12 characters long and can contain letters, numbers, hyphens (-) and underscores (_).')
            return None
        # A seat problem never stops you signing in: you just start without a seat.
        seat_note = None
        if seat in ('', '-'):
            seat = None
        elif not SEAT_RE.match(seat):
            seat_note = ('BAD_SEAT', '%s is not a valid seat number (for example, A01). Signed in without a seat.'
                         % clean(seat, 6))
            seat = None
        other = self.find_nick(nick)
        if other:
            if c.pcid and other.pcid == c.pcid:
                # Same PC signing in again, e.g. after a crash. Replace the old connection quietly.
                other.send('BYE', 'Disconnected: signed in again from this PC.')
                other.closed = True
                other.nick = None
                self.clients.discard(other)
            else:
                c.send('ERR', 'NICK_TAKEN', 'The nickname %s is already in use. Choose another nickname.' % other.nick)
                return None
        if seat:
            sitter = self.find_seat(seat)
            if sitter:
                seat_note = ('SEAT_TAKEN', 'Seat %s is taken by %s. Signed in without a seat.' % (seat, sitter.nick))
                seat = None
        c.nick, c.seat, c.status = nick, seat, 'here'
        self.gone.pop(nick.lower(), None)
        self.clear_held(seat, nick)
        if time.time() - self.arrived.get(nick.lower(), 0) > PARTY_HOURS * 3600:
            self.arrived[nick.lower()] = time.time()
        c.send('LOGGEDIN', nick, seat or '-', 1 if c.admin else 0)
        for t, n, text in self.history:
            c.send('HIST', t, n, text)
        self.send_schedule(c)
        self.send_vote(c)
        self.send_suggestions(c)
        self.send_food(c)
        self.send_map(c)
        if seat_note:
            c.send('ERR', *seat_note)
        if time.time() - self.updating.pop(nick.lower(), 0) < 300:
            self.broadcast('JOIN', nick, seat or '-', 'update', skip=c)
        else:
            self.broadcast('JOIN', nick, seat or '-', skip=c)
        if c.pcid:
            self.pcs[c.pcid] = {'nick': nick, 'seat': seat or '', 'last': datetime.datetime.now().isoformat(timespec='seconds')}
            self._save_pcs()
        log('%s signed in at seat %s from %s.' % (nick, seat or '-', c.addr))
        return None

    # ----- every few seconds, ask each LAN Center that can answer how long a round trip takes
    def published_client(self):
        """The version of the LAN Center on the file page, read again whenever the file changes."""
        p = FILES_DIR / 'client' / 'PalmsLAN.exe'
        try:
            st = p.stat()
            key = (st.st_mtime, st.st_size)
        except OSError:
            key = None
        if key != self._client_stat:
            self._client_stat = key
            self.client_ver = exe_version(p) if key else ''
            if self.client_ver:
                log('Palms LAN Center %s is on the file page. Older versions update themselves.' % self.client_ver)
        return self.client_ver

    def offer_update(self, c):
        """Tell a LAN Center that's older than the one on the file page, and where to get it."""
        if not self.file_page or not c.ver:
            return
        ver = self.published_client()
        if not ver or ver_tuple(ver) <= ver_tuple(c.ver):
            return
        try:
            here = c.writer.get_extra_info('sockname')[0]     # this server's address, as that PC sees it
        except Exception:
            return
        c.send('UPDATE', ver, 'http://%s:%s/' % (here, self.cfg.get('http_port', '8080')))

    async def pinger(self):
        self._announced = self.published_client()
        while True:
            await asyncio.sleep(PING_SECONDS)
            # a new LAN Center was just published: tell everyone connected
            ver = self.published_client()
            if ver != self._announced:
                self._announced = ver
                for c in list(self.clients):
                    self.offer_update(c)
            now = int(asyncio.get_running_loop().time() * 1000)
            for c in self.signed_in_clients():
                if 'ping' in c.caps:
                    c.send('SPING', now)

    # ----- "I'm here" shout so clients can find the server
    async def beacon(self):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.setblocking(False)
        message = ('PALMS|%s|%d|%s' % (self.name, self.tcp_port, VERSION)).encode('cp1252', 'replace')
        warned = False
        targets, looked = [], 0.0
        while True:
            # each network's own broadcast address (10.0.200.255): 255.255.255.255 needs a default
            # route, which a LAN without a router doesn't have. Looked up again every minute.
            if time.monotonic() - looked > 60:
                targets, looked = broadcast_addresses() or ['255.255.255.255'], time.monotonic()
            sent, err = 0, None
            for t in targets:
                try:
                    sock.sendto(message, (t, self.beacon_port))
                    sent += 1
                except OSError as e:
                    err = e
            if sent:
                warned = False
            elif not warned:
                log('Could not send the UDP beacon (%s). Clients can still connect by IP.' % err)
                warned = True
            await asyncio.sleep(self.beacon_seconds)

    def start_file_page(self):
        try:
            port = int(self.cfg.get('http_port', '8080'))
        except ValueError:
            port = 8080
        if port <= 0:
            return
        try:
            FILES_DIR.mkdir(exist_ok=True)
            (FILES_DIR / 'client').mkdir(exist_ok=True)
        except OSError as e:
            log('Could not make %s (%s). Uploads will fail until it exists.' % (FILES_DIR, e))
        for n in RETIRED:
            try:
                (FILES_DIR / 'client' / n).unlink()
                log('File page: removed client/%s (no longer used by Palms LAN Center).' % n)
            except FileNotFoundError:
                pass
            except OSError as e:
                log('File page: unable to remove client/%s (%s).' % (n, e.strerror))
        handler = make_files_handler(FILES_DIR, self.name)
        try:
            httpd = http.server.ThreadingHTTPServer(('0.0.0.0', port), handler)
        except OSError as e:
            log('Unable to start the file page on port %d (%s). The port may be in use.' % (port, e))
            return
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        self.file_page = True
        log('File page is running: http://<this server>:%d/  (files are in %s)' % (port, FILES_DIR))

    async def run(self):
        server = await asyncio.start_server(self.handle, '0.0.0.0', self.tcp_port, limit=MAX_LINE * 2)
        log('Palms Server %s started as %s on TCP port %d. Beacon on UDP port %d.'
            % (VERSION, self.name, self.tcp_port, self.beacon_port))
        self.tasks = [asyncio.create_task(self.beacon()), asyncio.create_task(self.pinger()),
                      asyncio.create_task(self.pak_watch())]
        self.start_file_page()
        async with server:
            await server.serve_forever()


# ---------------------------------------------------------------- systemd

UNIT = """[Unit]
Description=Palms Server
After=network-online.target
Wants=network-online.target

[Service]
User=palms
WorkingDirectory={here}
ExecStart={python} {script}
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
"""


def install_service():
    if os.geteuid() != 0:
        sys.exit('Run this as root: python3 %s --install-service' % pathlib.Path(__file__).name)
    load_config()
    DATA_DIR.mkdir(exist_ok=True)

    def run(*cmd, ok=(0,)):
        r = subprocess.run(cmd)
        if r.returncode not in ok:
            sys.exit('Command failed: %s' % ' '.join(cmd))

    if subprocess.run(['id', 'palms'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode != 0:
        run('useradd', '--system', '--home-dir', str(HERE), '--shell', '/usr/sbin/nologin', 'palms')
        print('Created system user: palms')
    run('chown', '-R', 'palms:palms', str(HERE))
    unit_path = pathlib.Path('/etc/systemd/system/palms.service')
    unit_path.write_text(UNIT.format(here=HERE, python=sys.executable, script=HERE / pathlib.Path(__file__).name))
    run('systemctl', 'daemon-reload')
    run('systemctl', 'enable', '--now', 'palms.service')
    run('systemctl', 'restart', 'palms.service')
    print('Palms Server installed and started. It starts automatically at boot.')
    print('  Status:  systemctl status palms')
    print('  Log:     journalctl -u palms -f   (Ctrl+C to exit)')


if __name__ == '__main__':
    if '--install-service' in sys.argv:
        install_service()
        sys.exit(0)
    cfg = load_config()
    try:
        asyncio.run(PalmsServer(cfg).run())
    except KeyboardInterrupt:
        log('Stopped.')
