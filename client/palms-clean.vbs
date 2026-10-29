' Palms LAN Center Cleanup - resets a test PC.
' Removes all games installed by Palms LAN Center, their registry settings and
' Windows Firewall exceptions, incomplete downloads, and the Palms LAN Center
' settings (the Palms Server then treats it as a new PC). The program is not removed.
' /quiet: nothing is asked or shown (the uninstaller of Palms LAN Center Setup runs it like this).
' Since 0.9.0 the settings and the list of games are in the registry, under
' HKEY_LOCAL_MACHINE\Software\Palms Arcade (palms.ini and gamepaks.ini until 0.8.1).
' Runs on Windows 98 (Windows Script Host 1.0) and later.
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

' 2. Each game it installed: from the registry, and from gamepaks.ini on a PC that never ran 0.9.0
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

' 4. The old settings files (0.8.1 and earlier), cached game icons and work files
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

' A LAN Center setting: the registry, or palms.ini on a PC that never ran 0.9.0
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
