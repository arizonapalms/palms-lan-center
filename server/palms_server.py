#!/usr/bin/env python3
"""Palms Server: the server every Palms LAN Center at the party connects to, for chat, the schedule,
votes, tournaments, the Game Library list and the file page."""

import asyncio
import base64
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
import subprocess
import sys

VERSION = '0.6.1'
HERE = pathlib.Path(__file__).resolve().parent
CONFIG_FILE = HERE / 'palms.ini'
DATA_DIR = HERE / 'data'
PCS_FILE = DATA_DIR / 'pcs.json'
SCHED_FILE = DATA_DIR / 'schedule.json'
TPL_FILE = DATA_DIR / 'templates.json'
MAPS_FILE = DATA_DIR / 'seatmaps.json'
FOOD_FILE = DATA_DIR / 'food.json'
CUPS_FILE = DATA_DIR / 'cups.json'
PROJ_FILE = DATA_DIR / 'projector.json'
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
    'SCHEDGAME|id|cartridge/game|server|how to connect   host only. The game an event is for, for the projector',
    'PROJ                     show the projector settings',
    'PROJSET|seconds|now,sched,cup,vote   host only. Seconds per slide, and the slides to show',
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
                'beacon_seconds': '3', 'admin_password': '', 'projector_port': '8080',
                'file_page': 'off', 'http_port': '8081', 'paks_url': 'http://10.0.200.73/'}
    changed = False
    # the file page used 8080, which is the projector's now
    if 'file_page' not in s and s.get('http_port') == '8080':
        s['http_port'] = '8081'
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


# ---------------------------------------------------------------- web pages
#
# The projector (projector_port, 8080): slides for the big screen, for Internet Explorer 5 and later.
#   /                    the slide show: two frames, the next slide loads in the hidden one
#   /slide?n=3&w=800     slide number 3, laid out for a screen 800 pixels wide
#   /gif/palm.gif        the palm and the event icons
# The file page (http_port, 8081), off unless file_page = on in palms.ini: a list of files
# everyone can download, and a form to upload more.

import html as _html
import http.server
import mimetypes
import socketserver
import threading
import time
import urllib.parse


class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


MAX_UPLOAD = 200 * 1024 * 1024
SAFE_NAME = re.compile(r'[^A-Za-z0-9._()\- ]')
DOWNLOAD_ONLY = ('.exe', '.ocx', '.dll', '.vbs', '.res', '.frx', '.zip', '.bat', '.reg', '.hta')


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


class QuietHandler(http.server.BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.0'

    def log_message(self, fmt, *args):
        pass

    def send_bytes(self, code, data, ctype, extra=None, cache=False):
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(data)))
        if cache:
            self.send_header('Cache-Control', 'max-age=86400')
        else:
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Pragma', 'no-cache')
            self.send_header('Expires', '0')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(data)

    def do_HEAD(self):
        self.do_GET()


def make_files_handler(root, server_name):

    class FilesHandler(QuietHandler):
        server_version = 'PalmsFiles/' + VERSION

        def resolve(self, url_path):
            """A real file under the files folder, or None. Never anything outside it."""
            rel = urllib.parse.unquote(url_path.split('?', 1)[0]).lstrip('/')
            try:
                p = (root / rel).resolve()
                p.relative_to(root.resolve())
            except (ValueError, OSError):
                return None
            return p if p.is_file() and not p.name.startswith('.') else None

        def do_GET(self):
            path = self.path.split('?', 1)[0]
            if path in ('/', '/index.html'):
                return self.page_with('')
            p = self.resolve(path)
            if not p:
                return self.send_bytes(404, b'File not found.', 'text/plain')
            ctype = mimetypes.guess_type(p.name)[0] or 'application/octet-stream'
            extra = {}
            if p.suffix.lower() in DOWNLOAD_ONLY:
                ctype = 'application/octet-stream'
                extra['Content-Disposition'] = 'attachment; filename="%s"' % p.name
            self.send_bytes(200, p.read_bytes(), ctype, extra)

        def do_POST(self):
            if self.path.split('?', 1)[0] != '/upload':
                return self.send_bytes(404, b'File not found.', 'text/plain')
            length = int(self.headers.get('Content-Length') or 0)
            m = re.search(r'boundary=("?)([^";]+)\1', self.headers.get('Content-Type', ''))
            if not m or length <= 0:
                return self.page_with('Select a file to upload.', bad=True)
            if length > MAX_UPLOAD:
                self.rfile.read(length)
                return self.page_with('File too large. Maximum file size: %s.' % nice_size(MAX_UPLOAD), bad=True)
            body = self.rfile.read(length)
            filename, data = '', None
            for part in body.split(b'--' + m.group(2).encode()):
                head, sep, content = part.partition(b'\r\n\r\n')
                if not sep:
                    continue
                if content.endswith(b'\r\n'):
                    content = content[:-2]
                fn = re.search(r'filename="([^"]*)"', head.decode('latin-1'))
                if fn is not None:
                    filename, data = fn.group(1), content
            name = safe_name(filename)
            if not name or not data:
                return self.page_with('Select a file to upload.', bad=True)
            tmp = root / ('.%s.part' % name)
            try:
                root.mkdir(parents=True, exist_ok=True)
                tmp.write_bytes(data)
                os.replace(tmp, root / name)
            except OSError as e:
                return self.page_with('Unable to save the file (%s). On the server, run: chown -R palms:palms %s'
                                      % (e.strerror, root), bad=True)
            log('File page: %s uploaded (%s) from %s.' % (name, nice_size(len(data)), self.client_address[0]))
            self.page_with('Uploaded %s.' % name)

        def page_with(self, msg, bad=False):
            items = sorted((p for p in root.glob('*') if p.is_file() and not p.name.startswith('.')),
                           key=lambda p: p.name.lower())
            rows = []
            for p in items:
                st = p.stat()
                rows.append('<tr><td><a href="/%s">%s</a></td><td class="r">%s</td><td>%s</td></tr>' % (
                    urllib.parse.quote(p.name), _html.escape(p.name), nice_size(st.st_size),
                    time.strftime('%a %d %b, %I:%M %p', time.localtime(st.st_mtime)).replace(' 0', ' ')))
            note = '<p class="%s">%s</p>' % ('bad' if bad else 'good', _html.escape(msg)) if msg else ''
            page = FILES_PAGE % {'name': _html.escape(server_name), 'note': note, 'limit': nice_size(MAX_UPLOAD),
                                 'rows': ''.join(rows) or '<tr><td colspan="3" class="none">No files.</td></tr>'}
            self.send_bytes(200, page.encode('utf-8'), 'text/html; charset=utf-8')

    return FilesHandler


FILES_PAGE = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Palms Server - Files</title>
<style type="text/css">
body { background: #008284; margin: 0; padding: 16px; font: 13px Tahoma, "MS Sans Serif", Arial, sans-serif; color: #000; }
.win { background: #C6C3C6; border: 2px outset #FFFFFF; max-width: 760px; margin: 0 auto; }
.title { background: #08599C; color: #FFF; font-weight: bold; padding: 4px 8px; }
.body { padding: 12px 14px; }
h2 { font-size: 15px; color: #103473; margin: 16px 0 6px; }
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
<h2>Files</h2>
<table><tr><th>Name</th><th>Size</th><th>Modified</th></tr>%(rows)s</table>
<h2>Upload File</h2>
<form method="post" action="/upload" enctype="multipart/form-data">
<table><tr><td>File:</td><td><input type="file" name="file" size="40"></td></tr>
<tr><td></td><td><input type="submit" value="Upload"> <span class="small">Maximum file size: %(limit)s. A file with the same name is replaced.</span></td></tr></table>
</form>
</div></div></body></html>
'''


# ---------------------------------------------------------------- the projector

PROJ_GIFS = {
    'palm': 'R0lGODlhIAAgAIMAAP/3xv/3re/ThJzb/5Tne/+2Id6iY8aiUudxEDmW5xiCIRhhtVpFGAg8CFokACkUACH5BAEAABAALAAAAAAgACAAQAj/ACEIHEiwoMGDBBsQUECgYQOFDxsqmPiwgMWLAQI4QEjwgYEHF0OKDImg5EaBAxIwAADggACXHh+8fHmAJYMEOHMmWMAT4UOOQBNKnBjzgdGJDAlUHFlSI9CYI6NaLIngJISUCQZo3apVp9evO3sGHcvxYYOBChWcFWhWaYOMGq0eNEu3bgOpFxE4ZbvQrNGPRtsiVSqSqtyCUPEyNYmQgeOYjiMzYFqA8UCskVse2Mx5c2SwYReg5Em6tOnTqMWSXc26dVCFSg02cECbNlDYdPtCiCjx7tSMtgtGHAxxIVLfUfUe5k0gMOyGbqUaRmuc6EcDH6snRU6yKtufAxNfTayb3DtC8Yq7H+6IHWR69U/bv69sGaVWxywBxMyf3zHX/1qhpJJmnBnVWWcMAIiTWFjlJNmDN4GWE4M7VbiAhRheqGGGYY2W2ocgihYQADs=',
    'trophy': 'R0lGODlhIAAgAIMAAP/////7xt7f3v/TjP/bKf/DKd6aQqWipc6aCMaKCKVlIYxZAFpBAEowADEwMUooACH5BAEAABAALAAAAAAgACAAQAj7ACEIHEiwoMGDCBMyCICgIYOEEAs2CECxIsUGESE4ACCgY8cDBxxMLFBggckECRo8WMmyZUaCDB6+nEmzoIObOHOOJHkyZcaNIIOCFBkApVGUKgcoWMpUwYOaMWtCjEmRAAKKUaUazJpVq1epO0suQPqTo0cBQ8P2xBhx41m0IcOSRPpggIG7d5c+1Zizb4O/gAO3HLz3q+GCXQ/HXIx1sUzDCxsieHwYwkKqlCFTTly5M8HAoNl+lWtUtNedJsf6bAugtevWREmKJcv6bNDYPFWbRujW423SdAcIHy58L1C0cOMWLZCAeXAFBqAzNQ5A6O+iR5831esZYUAAOw==',
    'crosshair': 'R0lGODlhIAAgAIIAAP+Sc+c0Ia0cEEoIAAAAAAAAAAAAAAAAACH5BAEAAAUALAAAAAAgACAAQAjnAAsIHEiwoMGDBwcAECBggEECBAwqZOiw4EQAFREinNhw40IBADAOGABxpMKPHTUOPMmwJUSUIDFq5JhRpcWPNQvQPAgxIU6bBE2WHAlU5ciQAF6GJFpUJ0qmBXoKPMpQpk+KPCNK/EkQotevWgsECDAQLNigXJtOTbvyo9WyYdcGwGqTqksCLSnmVMuXr1ACJvuuZMkQYsulez3m/WsS5durdOHerJp458O4bUHmtOx34WPOakGb/TpwLNzRcUE3VX0xsdTJKSc/Fvg6c+ytlDO+Pjr3tuKqIZXCTGyUo2G9gpMrFxwQADs=',
    'rocket': 'R0lGODlhIAAgAIMAAP//////5/f3///3pbXb/+fTc//bKbW6xnuq53uClBA0cyEkKQAAAAAAAAAAAAAAACH5BAEAAAwALAAAAAAgACAAQAj3ABkIHEiwoMGDCA0qAECgoUMEEBUkHLhAQIIEBgxcvJgxwQGPHTeGTLAgYYABKFOqVFmgQICJMGMmXFBRgM0DOHPq1GnTJs2YJ1eydCmTQYACA5AqTcp0qVKiRaNKFahgIYCrV6tOpUizq1evEyt+3Em2bE6PCQSUNPiVpsWNJNsWDSp0JVSgdfOivHvwaMu/gAML5ru1sGHDVRNLPMzAqkOHWQsvRECAsuXKCAAsjum1p+fPPruyrWm29M7QAiuKNDAWI+ucGVuPJEmw61u4uHNfFA00QIC2vl9upZu35VTixY3P1StUOUzkzJM6R+h3sHXAjKMGBAA7',
    'flag': 'R0lGODlhIAAgAIMAAP/7xpTne97f3s6aCDG2OWtpaxiCIVpBAAg8CCEgIQAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAAoALAAAAAAgACAAQAjgABUIPEDwgMCDCBMqXKgggYACCRAEIECxosWLFA0YQKDQIUSJGAlo1IigpEmJHBN6TMCwpcuGDyOmTFhyIkaSHWOCDMlT5MacEA+enPkSKMuiSAcCGGDwZc0AUAOUXLhSYtSrPX0SFVjVZtaLOFXq9HrT58iRW2F+DHD2bNqXK5MmjSu3aIK7R+u6PACgb1OnJ1smIJiXoVWLBqCm7Tr0aUi0Rnd+rQhZ7NrJiMMi7IqZsuaDnL8aEGl28ViebVObXju6dem3Uy2zrOn2LdyYem8HzS0YN2+GdH8bFU48aUAAOw==',
    'castle': 'R0lGODlhIAAgAIMAAP///8bDxpyanISChOc0IVJRUkJBQikoKSEgIQAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAAoALAAAAAAgACAAQAj/ABUIHEiwoMGDBhEAAIBA4QAEBA0YSLiwIYCHBBUGSJBgwEMAAUKKFOlRI8eSGUEG8PgxZAIBBQ4cSLCyZU2MAxtCVKAzogECQIP25Nkw5caOOHkurJgT5MmkSkOy1LhRgFUBNEuqnJqSpVaPL2XOnHrR606EBiWiRatwZ1ufBAq+VXo26s26auVuRdn0aEmdexkC9guVKtmRiLVK5SvQsOKqMWfWdFx4r82XkbPa5JpTp2edHEMn+Ex6renTqAVKDJqabmOGcIE2dQt79mu8E/XSrqtxIWMFeY36rkwyKQKJQ+1ONvr0rEPGJpHydip9KeIAS6P/pjpZ5fXDi6cXN+cu+q737ecvCxCL1ObyvuFVhpXZnjv68fLX0wf//vXK7gNUxd5hAaLn1Udg6TfWV2a15uBAAQEAOw==',
    'planet': 'R0lGODlhIAAgAIIAAP/7xt6m/86aCJxF3lpBAGMklCEIOQAAACH5BAEAAAcALAAAAAAgACAAQAjJAA8IHEiwoMGDCAsaMBAgwICHDws0XJjQIEOIGDNGLEAAgAACFRUyLECSJMWQKFOmvKhRo8mOBGKCPABz5sUCD2V2BBBzgEmVQIMKHUpwoVEDRAWybOnz5cyKS5lCLGlg51OBO0dKHSATAE+kB3d6HctTJ86SP5OqXcu2rdu3Qhc2nBvgZNyGWycCjcqUZN2yMRHy3Sqxqs6Dg6WS7PgxLIHELWOWDesRJGTJgMEa1Mm5q1eZZwtoRtg5ZgC0qO0SPaoaruvXagMCADs=',
    'tank': 'R0lGODlhIAAgAIIAAK3Dc5yanGuGMUJVGFJRUikoKRggCAAAACH5BAEAAAgALAAAAAAgACAAQAjWABEIHEiwoMGDCBMqXEjQgMOHECNGPHggAIEAGC1m1JjxYseNBQocYEiy5EADAAAIWMkyJcqUMGEaaAhggM2bOHPqHABgZsEDQIMKHUp0qMmjSJOSlKj0ZM2dNnuWRDlAQNWrVrNivcrTJ8GKG8OKHZtxpMCKITGGLKA2bYC1bdm+DWm2qd27ePPq3ct36cuYEgMnRKlVq83AEAc/hTrAK8O/MSNLdqmYseXDBhFr3uy1IgGPn0F/xhia9Oi5Acyi3ehRLWuQIEUORLu2tu3bt+v23W0wIAA7',
    'sword': 'R0lGODlhIAAgAIMAAP/////7xt7f3qWipd6iY86aCGs8EFpBADEwMSkUAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAAoALAAAAAAgACAAQAi0ABUIHEiwoMGDCBMeRABAgIABAxAoXNjwYcSJBxMQMHAgQIECAQ5gHElyIIKTACSWNFkRosqJHT+GdOhypcCOIRMYPMDTps+JL38KZAjxolAFDGka/ZnUYlCFOEMObbnUZ0yQB57a1GjAgE4FV0Xa7NjzqNmzTLUKZQggJVqkVNWubFrTLN2qTOOevSv3IM+yU5X2RXg1JN+jhQ/UNZsg8VmuXh2P5dn1602PYkuGfUswM+eAADs=',
    'bomb': 'R0lGODlhIAAgAIMAAP/bKZyanP+Sc6VpMec0Ia0cECkoKUoIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAAkALAAAAAAgACAAQAjcABMIHEiwoMGDCBMKHDBAocIDAggUmHgA4sQCFSNKxOgwAcSNGC1S/LjxQMeTCwcAWMkyIUQBMAVkjDmTpkMEOHPq3MkTgUORIwVczHiRI8qjCVYiNciwIUGlS1NGHfhxKNCMBEreDMC1q9evX30+jGhVaNCsFDtWDVoWrdGHFePKnTt3akeodg2ybGm3KUOBeKP+zYvQ6VS6cQcgjqtW49nHaX86DkmWMkiTCntq7pk5gIHPoEOLDi3W5WSskN+adova8tDGrNe2Vo3QYknZa2kfBEq5bFHMhI8GBAA7',
    'pizza': 'R0lGODlhIAAgAIMAAP/7xv/bKf/TjN6aQs6aCP+Sc6VlIVpBAEooAK0cEEoIAAAAAAAAAAAAAAAAAAAAACH5BAEAAAsALAAAAAAgACAAQAjhABcIHEiwoMGDCBEIMMCwocOHEBEgLHgAQICLGC8q2HhgoseBFQMQINDxI0GFAlKqXMmyZUqJEysqKECzgAKMGzeOLGmSYMiZNnf27HmA59CjSAUqHGCAqdOmUJ9KZQjzYMiMWDOOJDn0wMwEYG9qLBDWaFeLWLeaTSqTrAKuSRFWhBu3rt27CPLq3cu3r9+CSwcIHky4MOGGVQcqhMi4MdWDfiNL9ng1p9isIulSBvA1LM6ab9fG5JwzY2cFojcTwKy27tUAGy8KtTv361vNrknrTM0WQOu7FH3zpj0c+NGAADs=',
    'burger': 'R0lGODlhIAAgAIMAAP///8b3jP/TnN6iY+eiSv+Sc7VtITGGGFJRUms8EBA8COc0Ia0cEEooCEoIACkUACH5BAEAABAALAAAAAAgACAAQAj/ACEIHEiwoMGCCBAcXGjQQYEFECMuYEDRAcEGAgwQ0Mhxo0cDBgQ0kAixIsEHDwYkWMmyZUuUD0iaZMgQAYCbOHPiVFiwAUYBQIOKhOCQZMmZEBQoXRqgqdMAS6PKZGDxos+MILNq9dnAgdevX2mKPWiT59iFDguoXcu26lmCRY1SpNpTAIG7ePMSANrg4VS6AjFqHUw45EiJSCGgjMq4MUq5gAMD7Uj5Y0iRczO7fcu589iyngcmHE26tNm3plMnDB0X8mbOrae+tsqVq8DYiCMP/Kl3o+G+Ro9uxthbb1bgrndjLVz4cMS5w2tLn147d2SmTg9o337g6VLrblMOJVDpsvzK8TGfJ4bQAKb79zB9TpwP/eJy5oN9Zq5/kHrt0ACKFRAAOw==',
    'moon': 'R0lGODlhIAAgAIIAAP///+/v796iY5yanFJRUms8ECkUAAAAACH5BAEAAAcALAAAAAAgACAAQAjjAA8IHEhwAMEDAwweXMiwYEKHDSMKJACgosWLGAEECEARAAECDTtmHKmR40eJBxMqFPgQ5UIDAmIWmEmz5syYAgy4FEkS40aQBygCXUhxo9GjSDm6XNiSZdOlTFcifAr1oIGrWLNerRr0olatFX/u7DlSbNehBHmSDWtSaMiwGnsGKIkW6se7d7lGlUpV78CnfbmqhOiXsFOphac6RAwVK87HkLEuJZBVgM2aWesyVEvWaN6InOX+/Ojx7VqLnieWJnoatVm3rFuX5Fjx5GbZszWzRhp3ru+jujffnQ38c+LCAQEAOw==',
    'star': 'R0lGODlhIAAgAIIAAP///+/v79bT1oSChFJRUuc0ITk4OQAAACH5BAEAAAcALAAAAAAgACAAQAjuAA8IHEiwoMGDCA0AAEAAQICHDQEMmEixogABBhICKMDxIccCEiuKvJjxoEKGDiEuFDkSo8aFMGPKjHnRpcmZCz/iBFCzpMGTOUHupElSY82jSJP2TGigqdOnUKM2fUmgqtWhC5feZGi1alCcWn9uLNBV586wBU8S+EgWK8+iW7t6xYqWoEIBchsK4Ml3r1+4CAMLHky48Fa3Q32mRXzWpliUMCPSdbwY5cMAkocCrozT7MzNdnd6/kw5tNnRYEsPPMmWI2LQq2GyZQxbIFDGRBWHhrwwc2PdsS2rdFv7wF2lyJECty21eXPD0A8GBAA7',
    'door': 'R0lGODlhIAAgAIQAAP/75//7xv/ztefPc96iY86aCJyanJTnezG2OaVpMWs8EBiCIVJRUgg8CCkoKSkUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAABEALAAAAAAgACAAQAj/ACMIHEiwoMGDByEYYPCAQIIECiI+EEBRwIABDhxAQEhQIUOHECVWtIhRo8AGDRB6bPgwooKJFS9m3BihwYGUBlcC2MmzJ8+ZHAeuBOkSJkWZJoNGGNpSZEWbB6JGRUB1wc2OBhw0dClxgEWSQJUqNEC2rNmzZGkGHUqAwIO3I8EmXbuQZciXI5Gq5cj0rlG5e1XWBRmgQIC/ek9GxYn1Y1O8Tw8smEyZMmOBfYvmLUnzckEIoN+KFn2x9ADQgROiXs26deqcCxnInk17dsa5fOuKjns0sdjBD9u6jembruOmRkuHNW43ZPLiuY87J85Z6VLgyKmH9dy4eQDIJKvXPbwKW3rh5CjTozyAwKrnvtOP2qw8uf3ky/CdWlSfXvIC7vmBp1xSKAmmlUMKxBfedrkdyNVLpkFn3YTWBQQAOw==',
    'megaphone': 'R0lGODlhIAAgAIIAAP/7xv/bKf+Sc86aCFpBAOc0Ia0cEAAAACH5BAEAAAcALAAAAAAgACAAQAjWAA8IHEiwoMGDBgkACMCwoUMBBQwYcOhwgEUCBBVS3Mix4kWECi1+REgy48KOKBsOAIDRZEqIEjuKbDkw5ACOInPqHJmQZsmfPQkIHUq0aNGDGlHCnIiT54GkHZfKdAqVowCJTDfONHkzKtapA3zWPBng5kyjR0naFAu0rdu3cN8qBEC3rt27eNlW3SigrwCUK33upXg15lTBZK1+PexS6eKmiF8+1hq2ccrLVBPz9Qu48liLXmOaLUvaqcCQlHfuZOsyp9C4CQFcZA1b4WvYIGnj3v0zIAA7',
    'joystick': 'R0lGODlhIAAgAIMAAP/////7xpyanM6aCP+Sc4SChFJRUuc0IVpBAK0cEEoIACkoKSEgIQAAAAAAAAAAACH5BAEAAA4ALAAAAAAgACAAQAj2AB0IHEiwoMGDCA0qIHCgocMEEBUknFiQAYACDCgWbCBgIcQECgyIbECypEgDCAIMWIlAQAONDhQsZJiAAAGZMCd6/Agy50aSAoIKsIhRqEuSMIEuWMq0qVOkPgvijIpwocOHNSVSFejxQAKvXiNuNUg041iEZbc2kCnzJFChJE8imIvgpUGOTQ0s0Huy796/TO3mLFnyrGGqMm1OPWwV603Dja+G1bo18lWxZ3fy7Hk4JtvFnR2kDT1w9GHCZQn7fCugL+G+cAVvnKl4pICmLkWmDMC77sbbn/uKZMBAOEq6S2ULVOr0L9/mC6BOJIy3KXXSPgMCADs=',
    'calendar': 'R0lGODlhIAAgAIMAAP///9bT1qXL/5yanDl15xhFrec0ITk4OSEgIQgcSgAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAAoALAAAAAAgACAAQAj/ABUIHEiwoMGDBhMIKECAocOGEB9KbFigQAKCBwAAGMBR48aOGjkO8CgygMkDGD2qXMlypckAKA8iQEBwZk2aA20ezNiyZ8uXMQXyFEkSpIGjR0NyBJrypdOnUKMGLahTYFUFV68SVChAwMyuXb+CFRsWAdiLAxVWXMu2rdu2aIUqHTm3KN2PA5gO5Omzr0a9cv0CQJpUJWAFQ0HitWv38IHHkCNLnkwZYVacVjFj1WxZ8+WbCBOITjBzNGkEpkuPVi16qwACBGbChi17dm3aCCjGVaD2re/fu00LH058eErBfk8eJ1q3+WLle5Enh7lcMXPCBupCD8x88WDC2qlHI5fuczti593TdzTPMwAA9/Dfy49P/714uVHz67+PsL9/ggEBADs=',
    'disc': 'R0lGODlhIAAgAIMAAP/////7xv/bKcbDxv+Sc86aCISChOc0IVJRUq0cEEoIACEgIQAAAAAAAAAAAAAAACH5BAEAAAwALAAAAAAgACAAQAj/ABkIHEiwoMGDBxcAGMCwocOHEA0YWGBQIUMFBA4cCCCgo8cCBQIcSJBAwQCJFAtaHKCgpcuXMF2enFhxpYGGNxnmPInTAIAFKQmuxEigqNGjSE2irLiQZUaNAUBKBamRpFKaKgFI3Iqgq9evXbduDaoSqNmzaNOSFdqUp06ccN8uzQqxrt25bC8+hRqgb1+NI0vOXMtg6NOoHjtKDXyVsGGSkCNLFox34Eq7mHViFaq2s2fCCEOLTghAa2nTpX2mPu0TdOG2O2PHndma6WzZbzVXFng5c92xtlnGHP4SOF2ngJMrh6zAeF7kUBODFMncueWmGKMnVlyg+u7XeqGCNxRQgDz57lat88ZOIMHIjVO7By6pHrzTyfjTf+/t+2F9hTcFCNaACPD0n1ZiJajggq6N5iBBAQEAOw==',
    'clock': 'R0lGODlhIAAgAIIAAP///9bT1qXL/4SChOc0IRhFrQgcSgAAACH5BAEAAAgALAAAAAAgACAAQAjyABEIHEiwoMGDBw0IMACgocOGBw48fKjQgMGKEzMCiBgxY0WEChcaCKCxoQGFBSwiJBgyJcOJJwXIVJlwYUmIEm9+ZDnzpk4BKVdePEn0pNCjCGLKXDqTJlKULx8OGOBxJkibJTtqDLBzIEYAUzMSIJAxLICuSXv6dMjRIdSLaieahQnUKciQIovKdIm0r9++RYv+LajUpd6mfqFG1ajULuG6GrXSXVgzZdacVSnztHz5p2aBXzuX5Po57cu5DsdKBkuVtNPQPjliBuD68eK1dB2bxj2x9lWsUqk+JB30KN6TtJNzbalbqFKmdZsPnk59cEAAOw==',
}

PROJ_SLIDES = ('now', 'sched', 'cup', 'vote')
VOTE_SHOW_MINUTES = 10       # a closed vote stays on the projector this long
CUP_SHOW_MINUTES = 60        # and a finished tournament this long
PROJ_BG = '#0B1416'

PROJ_PAGE = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>%(name)s - Projector</title>
<style type="text/css">
body { margin: 0; background: #0B1416; overflow: hidden; }
iframe { position: absolute; left: 0; top: 0; border: 0; }
</style>
<script type="text/javascript">
var n = 0, cur = 0, secs = 15;
function fr(i) { return document.getElementById('f' + i); }
function next() {
  var w = document.body.clientWidth, h = document.body.clientHeight, f = fr(1 - cur);
  f.style.width = w + 'px';
  f.style.height = h + 'px';
  f.src = '/slide?n=' + n + '&w=' + w + '&h=' + h + '&t=' + new Date().getTime();
  n++;
  setTimeout('show(' + (1 - cur) + ')', 1500);
}
function show(i) {
  fr(i).style.visibility = 'visible';
  fr(1 - i).style.visibility = 'hidden';
  cur = i;
  setTimeout('next()', secs * 1000 - 1500);
}
</script></head>
<body onload="next()" scroll="no">
<iframe id="f0" frameborder="0" scrolling="no" style="visibility: hidden" src="about:blank"></iframe>
<iframe id="f1" frameborder="0" scrolling="no" style="visibility: hidden" src="about:blank"></iframe>
<noscript><p style="color: #FFFFFF; font: 20px Arial">The projector needs JavaScript. In Internet Explorer: Tools &gt; Internet Options &gt; Security &gt; Custom Level &gt; Active scripting &gt; Enable.</p></noscript>
</body></html>
'''


def make_projector_handler(srv):

    class ProjectorHandler(QuietHandler):
        server_version = 'PalmsProjector/' + VERSION

        def do_GET(self):
            path, _, query = self.path.partition('?')
            q = urllib.parse.parse_qs(query)
            if path in ('/', '/index.html'):
                page = PROJ_PAGE % {'name': _html.escape(srv.name)}
                return self.send_bytes(200, page.encode('cp1252', 'replace'), 'text/html; charset=windows-1252')
            if path == '/slide':
                def num(k, d):
                    try:
                        return int(q.get(k, [d])[0])
                    except ValueError:
                        return d
                try:
                    page = srv.projector_slide(num('n', 0), num('w', 800), num('h', 600))
                except Exception as e:
                    log('Projector: unable to make a slide (%s).' % e)
                    page = slide_page(srv, 1.0, 'clock', 'PALMS ARCADE', '', '<br>', '')
                return self.send_bytes(200, page.encode('cp1252', 'replace'), 'text/html; charset=windows-1252')
            if path.startswith('/gif/') and path.endswith('.gif'):
                b = PROJ_GIFS.get(path[5:-4])
                if b:
                    return self.send_bytes(200, base64.b64decode(b), 'image/gif', cache=True)
            self.send_bytes(404, b'Not found.', 'text/plain')

    return ProjectorHandler


def pesc(s):
    """HTML-safe text for Internet Explorer 5 (no &#x27;)."""
    return _html.escape(str(s), quote=False).replace('"', '&quot;')


def px(s, n):
    return max(1, int(round(n * s)))


def slide_page(srv, s, icon, title, right, body, foot):
    """One slide: a header bar with the palm, a title and the time, the body, and a line at the bottom."""
    e = pesc
    return ('<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">\n<html><head>'
            '<meta http-equiv="Content-Type" content="text/html; charset=windows-1252"><title>%(title)s</title>'
            '<style type="text/css">\n'
            'body { margin: 0; background: %(bg)s; color: #FFFFFF; font-family: Arial, Helvetica, sans-serif; }\n'
            'td { font-family: Arial, Helvetica, sans-serif; color: #FFFFFF; }\n'
            '.hd { background: #31616B; }\n'
            '.t { font: bold %(t)dpx Arial, Helvetica, sans-serif; color: #FFFFFF; letter-spacing: 1px; }\n'
            '.r { font: bold %(r)dpx Arial, Helvetica, sans-serif; color: #BFD6DA; }\n'
            '.big { font: bold %(big)dpx Arial, Helvetica, sans-serif; color: #FFDB29; }\n'
            '.mid { font: bold %(mid)dpx Arial, Helvetica, sans-serif; color: #FFFFFF; }\n'
            '.txt { font: %(txt)dpx Arial, Helvetica, sans-serif; color: #FFFFFF; }\n'
            '.lbl { font: bold %(lbl)dpx Arial, Helvetica, sans-serif; color: #8CAAB5; letter-spacing: 2px; }\n'
            '.addr { font: bold %(addr)dpx "Courier New", Courier, monospace; color: #FFFFFF; }\n'
            '.dim { color: #6B8A92; }\n'
            '.now { color: #FFDB29; }\n'
            '.ft { font: bold %(ft)dpx Arial, Helvetica, sans-serif; color: #BFD6DA; }\n'
            '</style><script type="text/javascript">try { parent.secs = %(secs)d; } catch (x) { }</script></head>'
            '<body scroll="no"><table width="100%%" cellspacing="0" cellpadding="0" class="hd"><tr>'
            '<td width="%(hw)d" align="center"><img src="/gif/palm.gif" width="%(pal)d" height="%(pal)d" alt=""></td>'
            '<td width="%(iw)d"><img src="/gif/%(icon)s.gif" width="%(ic)d" height="%(ic)d" alt=""></td>'
            '<td class="t" height="%(hh)d">%(title)s</td><td class="r" align="right">%(right)s&nbsp;&nbsp;</td></tr></table>'
            '<table width="100%%" cellspacing="0" cellpadding="0"><tr><td bgcolor="#F78E21" height="%(st)d"></td></tr></table>'
            '<table width="100%%" height="%(bh)d" cellspacing="0" cellpadding="%(pad)d"><tr><td valign="top">%(body)s</td></tr></table>'
            '<table width="100%%" cellspacing="0" cellpadding="%(pad)d"><tr><td class="ft">%(foot)s</td></tr></table>'
            '</body></html>') % {
        'bg': PROJ_BG, 'title': e(title), 'right': e(right), 'body': body, 'foot': foot, 'icon': icon,
        't': px(s, 34), 'r': px(s, 24), 'big': px(s, 44), 'mid': px(s, 30), 'txt': px(s, 24), 'lbl': px(s, 16),
        'addr': px(s, 40), 'ft': px(s, 22), 'secs': srv.proj['seconds'], 'hh': px(s, 64), 'hw': px(s, 76),
        'pal': px(s, 48), 'iw': px(s, 44), 'ic': px(s, 32), 'st': px(s, 4), 'pad': px(s, 14),
        'bh': px(s, 600) - px(s, 64) - px(s, 4) - px(s, 22) - 3 * px(s, 14)}


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
        self.loop = None
        self.updating = {}        # nick (lower case) -> when they closed to install an update
        DATA_DIR.mkdir(exist_ok=True)
        self.pcs = self._load_pcs()
        self.schedule = self._load_schedule()
        self.templates = self._load_templates()
        self.vote = None          # the current or last vote, see start_vote()
        self.suggest = []         # games suggested for the next vote: {'name', 'by': [nick, ...]}
        self.food = self._load_food()
        self.cups = self._load_cups()
        self.proj = self._load_proj()
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
        for e in events:
            if e.get('game') or e.get('server') or e.get('connect'):
                c.send('SCHEDGAME', e['id'], e.get('game', ''), e.get('server', ''), e.get('connect', ''))
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
            if champ:
                t.setdefault('done_at', time.time())
            else:
                t.pop('done_at', None)
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
        v['closed_at'] = time.time()
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
                text = await asyncio.get_running_loop().run_in_executor(None, fetch_text, url + 'catalog.txt')
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
        if cmd == 'PROJ':
            self.send_proj(c)
            return None
        if cmd == 'PROJSET':
            if not c.admin:
                c.send('ERR', 'NOT_ADMIN', 'Host mode is required to change the projector. Use HOST|password first.')
                return None
            try:
                secs = int(arg(1))
            except ValueError:
                secs = 0
            if not 5 <= secs <= 300:
                c.send('ERR', 'BAD_SECONDS', 'Each slide can show for 5 to 300 seconds.')
                return None
            self.proj = {'seconds': secs, 'slides': [x for x in PROJ_SLIDES if x in arg(2).lower().split(',')]}
            self.proj_changed()
            log('%s changed the projector: %d seconds, %s' % (c.nick, secs, ','.join(self.proj['slides']) or 'no slides'))
            return None
        if cmd == 'SCHEDGAME':
            if not c.admin:
                c.send('ERR', 'NOT_ADMIN', 'Host mode is required to change What\'s On. Use HOST|password first.')
                return None
            e = next((e for e in self.schedule if str(e['id']) == arg(1)), None)
            if not e:
                c.send('ERR', 'NOT_FOUND', 'Event %s not found. Type SCHED to list events.' % clean(arg(1), 10))
                return None
            game = arg(2).strip().lower()
            if not re.match(r'^[a-z0-9-]{1,24}(/[a-z0-9-]{1,24})?$', game):
                game = ''
            for k, v in (('game', game), ('server', clean(arg(3), 40).replace('|', '/')),
                         ('connect', clean(arg(4, rest=True), 120).replace('|', '/'))):
                if v:
                    e[k] = v
                else:
                    e.pop(k, None)
            self.schedule_changed()
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
        self.send_proj(c)
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
    # ----- the projector: its settings, and the slides (made for the web page on projector_port)
    def _load_proj(self):
        p = {'seconds': 15, 'slides': list(PROJ_SLIDES)}
        try:
            d = json.loads(PROJ_FILE.read_text(encoding='utf-8'))
            if 5 <= int(d.get('seconds', 15)) <= 300:
                p['seconds'] = int(d['seconds'])
            p['slides'] = [x for x in d.get('slides', []) if x in PROJ_SLIDES]
        except FileNotFoundError:
            pass
        except Exception as e:
            log('Could not read %s (%s). Using the standard projector settings.' % (PROJ_FILE, e))
        return p

    def send_proj(self, c):
        c.send('PROJINFO', self.proj['seconds'], ','.join(self.proj['slides']), self.cfg.get('projector_port', '8080'))

    def proj_changed(self):
        tmp = PROJ_FILE.with_suffix('.tmp')
        tmp.write_text(json.dumps(self.proj, indent=1), encoding='utf-8')
        os.replace(tmp, PROJ_FILE)
        for c in self.signed_in_clients():
            self.send_proj(c)

    def catalog_game(self, key):
        """A game from the Cartridge Server's list, by its key (pak/game): name, icon and screenshot addresses."""
        pak_id, _, game_id = key.partition('/')
        out = None
        for line in self.paks:
            f = line.split('|')
            if f[0] == 'PAK' and len(f) > 2 and f[1] == pak_id:
                out = {'name': f[2], 'ver': f[3] if len(f) > 3 else '',
                       'icon': f[18] if len(f) > 18 else '', 'shot': f[21] if len(f) > 21 else ''}
            elif out and f[0] == 'GAME' and len(f) > 3 and f[1] == pak_id and f[2] == game_id:
                out['name'] = f[3]
                if len(f) > 10 and f[10]:
                    out['icon'] = f[10]
        if out:
            for k in ('icon', 'shot'):
                v = out[k]
                if v and not v.startswith('http'):
                    v = self.paks_base + v
                out[k] = v
            if out['icon'].lower().endswith('.bmp'):
                out['icon'] = out['icon'][:-4] + '.gif'
        return out

    def proj_events(self):
        """Today's events in order, the one on now (the last one started) and the next one."""
        events = sorted(self.schedule, key=lambda e: (party_minutes(e['time']), e['id']))
        mins = party_minutes(now_hm())
        cur = None
        for e in events:
            if party_minutes(e['time']) <= mins:
                cur = e
        nxt = next((e for e in events if party_minutes(e['time']) > mins), None)
        return events, cur, nxt

    def proj_list(self):
        """The slides there's something to show on, in order."""
        on = self.proj['slides']
        out = []
        events, cur, nxt = self.proj_events()
        if 'now' in on and (cur or nxt):
            out.append(('now', None))
        if 'sched' in on and events:
            out.append(('sched', None))
        if 'cup' in on:
            for t in self.cups:
                if t['state'] in ('signup', 'drawn', 'running'):
                    out.append(('cup', t))
                elif t['state'] == 'done' and time.time() - t.get('done_at', 0) < CUP_SHOW_MINUTES * 60:
                    out.append(('cup', t))
        v = self.vote
        if 'vote' in on and v:
            if v['state'] == 'open' or time.time() - v.get('closed_at', 0) < VOTE_SHOW_MINUTES * 60:
                out.append(('vote', None))
        return out

    def projector_slide(self, n, w, h):
        w, h = max(320, min(w, 4096)), max(240, min(h, 4096))
        s = min(w / 800.0, h / 600.0)
        slides = self.proj_list()
        clock = ampm(now_hm())
        if not slides:
            body = ('<table width="100%%" height="%d"><tr><td align="center" valign="middle">'
                    '<img src="/gif/palm.gif" width="%d" height="%d" alt=""><br><br>'
                    '<span class="big">%s</span><br><span class="txt dim">Palms Arcade LAN Party</span></td></tr></table>'
                    % (px(s, 380), px(s, 128), px(s, 128), pesc(self.name)))
            return slide_page(self, s, 'palm', 'WELCOME', clock, body, '')
        kind, arg = slides[n % len(slides)]
        return getattr(self, 'slide_' + kind)(s, clock, arg)

    def up_next(self, nxt):
        if not nxt:
            return ''
        return 'UP NEXT: %s &nbsp;%s' % (ampm(nxt['time']), pesc(nxt['name']))

    def slide_now(self, s, clock, arg):
        e = pesc
        events, cur, nxt = self.proj_events()
        ev, title = (cur, 'NOW PLAYING') if cur else (nxt, 'COMING UP')
        game = self.catalog_game(ev.get('game', '')) if ev.get('game') else None
        icon_img = ''
        if game and game['icon']:
            icon_img = '<img src="%s" width="%d" height="%d" alt="" align="absmiddle">&nbsp;' % (e(game['icon']), px(s, 64), px(s, 64))
        else:
            icon_img = '<img src="/gif/%s.gif" width="%d" height="%d" alt="" align="absmiddle">&nbsp;' % (ev['icon'], px(s, 64), px(s, 64))
        rows = ['<span class="big">%s%s</span>' % (icon_img, e(game['name'] if game else ev['name']))]
        if game and game['name'] != ev['name']:
            rows.append('<span class="mid">%s</span>' % e(ev['name']))
        rows.append('<span class="txt dim">%s %s</span>' % ('Started' if cur else 'Starts', ampm(ev['time'])))
        if ev.get('server'):
            rows.append('<br><span class="lbl">SERVER</span><br><span class="addr">%s</span>' % e(ev['server']))
        if ev.get('connect'):
            rows.append('<br><span class="lbl">HOW TO CONNECT</span><br><span class="txt">%s</span>' % e(ev['connect']))
        elif ev.get('note') and not ev.get('server'):
            rows.append('<br><span class="txt">%s</span>' % '<br>'.join(e(x) for x in ev['note'].split('~')[:4]))
        left = '<br>'.join(rows)
        shot = ''
        if game and game['shot']:
            shot = ('<td width="%d" valign="top" align="right"><table cellspacing="0" cellpadding="%d" bgcolor="#F78E21"><tr><td>'
                    '<img src="%s" width="%d" height="%d" alt=""></td></tr></table></td>'
                    % (px(s, 336), px(s, 3), e(game['shot']), px(s, 320), px(s, 240)))
        body = '<table width="100%%" cellspacing="0" cellpadding="0"><tr><td valign="top">%s</td>%s</tr></table>' % (left, shot)
        return slide_page(self, s, 'joystick', title, clock, body, self.up_next(nxt if cur else None))

    def slide_sched(self, s, clock, arg):
        e = pesc
        events, cur, nxt = self.proj_events()
        start = 0
        if cur:
            start = max(0, events.index(cur) - 1)
        start = max(0, min(start, len(events) - 7))
        rows = []
        for ev in events[start:start + 7]:
            cls = 'now' if ev is cur else ('dim' if cur and party_minutes(ev['time']) < party_minutes(cur['time']) else '')
            extra = ''
            if ev.get('server'):
                extra = ' <span class="txt dim">%s</span>' % e(ev['server'])
            rows.append('<tr><td class="mid %s" align="right" nowrap width="%d">%s</td>'
                        '<td width="%d" align="center"><img src="/gif/%s.gif" width="%d" height="%d" alt=""></td>'
                        '<td class="mid %s">%s%s%s</td></tr>'
                        % (cls, px(s, 170), ampm(ev['time']), px(s, 56), ev['icon'], px(s, 40), px(s, 40), cls,
                           '&raquo; ' if ev is cur else '', e(ev['name']), extra))
        body = '<table width="100%%" cellspacing="0" cellpadding="%d">%s</table>' % (px(s, 8), ''.join(rows))
        more = len(events) - (start + 7)
        foot = '%d more later today' % more if more > 0 else ''
        return slide_page(self, s, 'calendar', "TODAY'S SCHEDULE", clock, body, foot)

    def slide_cup(self, s, clock, t):
        e = pesc
        title = 'TOURNAMENT: %s' % t['name'].upper()
        sub = ' &nbsp;'.join(x for x in (e(t.get('game') or ''), ('starts ' + ampm(t['time'])) if TIME_RE.match(t.get('time') or '') else '') if x)
        if t['state'] == 'signup':
            players = t['players']
            cells = ''.join('<td class="mid" width="25%%">%s</td>%s' % (e(p), '</tr><tr>' if i % 4 == 3 else '')
                            for i, p in enumerate(players[:32]))
            body = ('<span class="big">SIGN-UPS OPEN</span><br><span class="txt dim">%s</span><br><br>'
                    '<span class="lbl">%d PLAYER%s SO FAR</span><table width="100%%" cellpadding="%d"><tr>%s</tr></table>'
                    '<span class="txt">Sign up in Palms LAN Center: Tournaments &gt; Sign Up.</span>'
                    % (sub, len(players), '' if len(players) == 1 else 'S', px(s, 4), cells))
            return slide_page(self, s, 'trophy', title, clock, body, '')
        rounds = self.cup_bracket(t)
        champ = t.get('champion') if t['state'] == 'done' else ''
        # rounds that fit: the bracket from the first round with 8 matches or fewer
        r0 = 0
        while r0 < len(rounds) - 1 and len(rounds[r0]) > 8:
            r0 += 1
        playing = next((r for r, ms in enumerate(rounds) if any(m['a'] and m['b'] and not m['win'] for m in ms)), None)
        if playing is not None and playing < r0:
            r0 = playing
            shown = [rounds[r0]]
        else:
            shown = rounds[r0:]

        def who(p, won, lost):
            if p is None:
                return '<span class="dim">bye</span>'
            if not p:
                return '<span class="dim">-</span>'
            return '<span class="%s">%s</span>' % ('now' if won else ('dim' if lost else ''), e(p))

        cols = []
        fs = px(s, 22 if len(shown[0]) <= 4 else (18 if len(shown[0]) <= 8 else 15))
        for r, ms in zip(range(r0, r0 + len(shown)), shown):
            cells = []
            for m in ms[:16]:
                wa, wb = m['win'] == 'a', m['win'] == 'b'
                score = (' <span class="dim">%s</span>' % e(m['score'])) if m['score'] else ''
                cells.append('<table width="100%%" cellspacing="0" cellpadding="%d" style="border: 1px solid #31616B; '
                             'margin: %dpx 0"><tr><td style="font: bold %dpx Arial">%s</td></tr>'
                             '<tr><td style="font: bold %dpx Arial">%s%s</td></tr></table>'
                             % (px(s, 3), px(s, 3), fs, who(m['a'], wa, wb), fs, who(m['b'], wb, wa), score))
            cols.append('<td valign="middle" width="%d%%"><span class="lbl">%s</span>%s</td>'
                        % (100 // len(shown), self.round_name(r, len(rounds)).upper(), ''.join(cells)))
        body = ''
        if champ:
            body = '<span class="big">CHAMPION: %s</span><br>' % e(champ)
        elif sub:
            body = '<span class="txt dim">%s</span><br>' % sub
        body += '<table width="100%%" cellspacing="%d" cellpadding="0"><tr>%s</tr></table>' % (px(s, 6), ''.join(cols))
        return slide_page(self, s, 'trophy', title, clock, body, '')

    def slide_vote(self, s, clock, arg):
        e = pesc
        v = self.vote
        counts = self.vote_counts()
        total = sum(counts)
        best = max(counts) if counts else 0
        rows = []
        for i, name in enumerate(v['options']):
            won = v['state'] == 'closed' and v.get('winner') == i + 1
            pct = (100 * counts[i] // total) if total else 0
            bar = ('<table width="%d%%" cellspacing="0" cellpadding="0"><tr><td bgcolor="%s" height="%d"></td></tr></table>'
                   % (max(pct, 1), '#FFDB29' if won or (counts[i] == best and best) else '#31616B', px(s, 22)))
            rows.append('<tr><td class="mid %s" width="40%%">%s</td><td>%s</td><td class="mid" width="%d" align="right">%d</td></tr>'
                        % ('now' if won else '', e(name), bar, px(s, 60), counts[i]))
        if v['state'] == 'open':
            left = max(0, int(v['deadline'] - self.loop.time()))
            status = 'Vote in Palms LAN Center: What\'s On &gt; Vote. Closes in %d:%02d.' % (left // 60, left % 60)
        elif v.get('winner'):
            status = '%s won with %d of %d votes.' % (e(v['options'][v['winner'] - 1]), counts[v['winner'] - 1], total)
        else:
            status = 'Voting closed.'
        body = ('<span class="big">%s</span><br><br><table width="100%%" cellspacing="0" cellpadding="%d">%s</table>'
                % (e(v['question']), px(s, 8), ''.join(rows)))
        return slide_page(self, s, 'star', 'NEXT GAME VOTE' if v['state'] == 'open' else 'VOTE RESULT', clock, body,
                          '%s &nbsp;%d vote%s' % (status, total, '' if total == 1 else 's'))

    def start_projector(self):
        try:
            port = int(self.cfg.get('projector_port', '8080'))
        except ValueError:
            port = 8080
        if port <= 0:
            return
        try:
            httpd = ThreadingHTTPServer(('0.0.0.0', port), make_projector_handler(self))
        except OSError as e:
            log('Unable to start the projector page on port %d (%s). The port may be in use.' % (port, e))
            return
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        log('Projector page is running: http://<this server>:%d/' % port)

    async def pinger(self):
        while True:
            await asyncio.sleep(PING_SECONDS)
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
        if self.cfg.get('file_page', 'off').strip().lower() not in ('on', 'yes', '1', 'true'):
            return
        try:
            port = int(self.cfg.get('http_port', '8081'))
        except ValueError:
            port = 8081
        if port <= 0:
            return
        try:
            FILES_DIR.mkdir(exist_ok=True)
        except OSError as e:
            log('Could not make %s (%s). Uploads will fail until it exists.' % (FILES_DIR, e))
        try:
            httpd = ThreadingHTTPServer(('0.0.0.0', port), make_files_handler(FILES_DIR, self.name))
        except OSError as e:
            log('Unable to start the file page on port %d (%s). The port may be in use.' % (port, e))
            return
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        log('File page is running: http://<this server>:%d/  (files are in %s)' % (port, FILES_DIR))

    async def run(self):
        server = await asyncio.start_server(self.handle, '0.0.0.0', self.tcp_port, limit=MAX_LINE * 2)
        log('Palms Server %s started as %s on TCP port %d. Beacon on UDP port %d.'
            % (VERSION, self.name, self.tcp_port, self.beacon_port))
        self.tasks = [asyncio.create_task(self.beacon()), asyncio.create_task(self.pinger()),
                      asyncio.create_task(self.pak_watch())]
        self.loop = asyncio.get_running_loop()
        self.start_projector()
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
