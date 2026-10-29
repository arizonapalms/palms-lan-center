Attribute VB_Name = "modReg"
' Palms LAN Center - Reads and writes the settings and installed games in the registry (HKLM\Software\Palms Arcade).
Option Explicit

Public Const REG_LAN As String = "Software\Palms Arcade\LAN Center"
Public Const REG_GAMES As String = "Software\Palms Arcade\Games"
Public Const REG_UNINSTALL As String = "Software\Microsoft\Windows\CurrentVersion\Uninstall"

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const KEY_READ As Long = &H20019
Private Const KEY_ALL_ACCESS As Long = &HF003F
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_MORE_DATA As Long = 234

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByVal lpcbClass As Long, ByVal lpftLastWriteTime As Long) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private mBroken As Boolean      ' the registry refused a write: said once, not every time

' ================================================================ reading and writing

' A value as text, or def when it isn't there
Public Function RegGet(ByVal path As String, ByVal valueName As String, ByVal def As String) As String
    Dim h As Long, kind As Long, size As Long, buf As String, n As Long
    RegGet = def
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, path, 0, KEY_READ, h) <> ERROR_SUCCESS Then Exit Function
    size = 0
    If RegQueryValueExStr(h, valueName, 0, kind, vbNullString, size) = ERROR_SUCCESS Then
        If kind = REG_DWORD Then
            n = 0
            size = 4
            If RegQueryValueExLong(h, valueName, 0, kind, n, size) = ERROR_SUCCESS Then RegGet = CStr(n)
        ElseIf size <= 1 Then
            RegGet = ""
        Else
            buf = String$(size, vbNullChar)
            If RegQueryValueExStr(h, valueName, 0, kind, buf, size) = ERROR_SUCCESS Then
                n = InStr(buf, vbNullChar)
                If n > 0 Then buf = Left$(buf, n - 1)
                RegGet = buf
            End If
        End If
    End If
    RegCloseKey h
End Function

' Is the value there at all?
Public Function RegHas(ByVal path As String, ByVal valueName As String) As Boolean
    Dim h As Long, kind As Long, size As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, path, 0, KEY_READ, h) <> ERROR_SUCCESS Then Exit Function
    RegHas = (RegQueryValueExStr(h, valueName, 0, kind, vbNullString, size) = ERROR_SUCCESS)
    RegCloseKey h
End Function

Public Function RegKeyThere(ByVal path As String) As Boolean
    Dim h As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, path, 0, KEY_READ, h) = ERROR_SUCCESS Then
        RegKeyThere = True
        RegCloseKey h
    End If
End Function

' Sets a value as text, making the key if it isn't there. False if Windows said no.
Public Function RegPut(ByVal path As String, ByVal valueName As String, ByVal value As String) As Boolean
    Dim h As Long, rc As Long
    If Not OpenForWriting(path, h) Then Exit Function
    ' An empty value goes as a lone zero. A string that was never set (vbNullString) would reach
    ' Windows as no data at all, which it refuses.
    If Len(value) = 0 Then
        rc = RegSetValueExStr(h, valueName, 0, REG_SZ, vbNullChar, 1)
    Else
        rc = RegSetValueExStr(h, valueName, 0, REG_SZ, value, Len(value) + 1)
    End If
    RegCloseKey h
    RegPut = (rc = ERROR_SUCCESS)
    If Not RegPut Then Refused path, rc
End Function

Public Function RegPutNumber(ByVal path As String, ByVal valueName As String, ByVal value As Long) As Boolean
    Dim h As Long, rc As Long
    If Not OpenForWriting(path, h) Then Exit Function
    rc = RegSetValueExLong(h, valueName, 0, REG_DWORD, value, 4)
    RegCloseKey h
    RegPutNumber = (rc = ERROR_SUCCESS)
    If Not RegPutNumber Then Refused path, rc
End Function

Public Sub RegDrop(ByVal path As String, ByVal valueName As String)
    Dim h As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, path, 0, KEY_ALL_ACCESS, h) <> ERROR_SUCCESS Then Exit Sub
    RegDeleteValue h, valueName
    RegCloseKey h
End Sub

' Deletes a key and everything in it. (Windows 98 would do that by itself; XP only deletes empty keys.)
Public Sub RegDropKey(ByVal path As String)
    Dim h As Long, names() As String, n As Long, i As Long, nm As String, size As Long, parent As String, p As Long
    If Len(path) = 0 Then Exit Sub
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, path, 0, KEY_ALL_ACCESS, h) <> ERROR_SUCCESS Then Exit Sub
    ReDim names(0 To 7)
    n = 0
    Do
        nm = String$(260, vbNullChar)
        size = 260
        If RegEnumKeyEx(h, n, nm, size, 0, vbNullString, 0, 0) <> ERROR_SUCCESS Then Exit Do
        If n > UBound(names) Then ReDim Preserve names(0 To n + 8)
        names(n) = Left$(nm, size)
        n = n + 1
    Loop
    RegCloseKey h
    For i = 0 To n - 1
        RegDropKey path & "\" & names(i)
    Next
    p = InStrRev(path, "\")
    If p = 0 Then Exit Sub
    parent = Left$(path, p - 1)
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, parent, 0, KEY_ALL_ACCESS, h) <> ERROR_SUCCESS Then Exit Sub
    RegDeleteKey h, Mid$(path, p + 1)
    RegCloseKey h
End Sub

Private Function OpenForWriting(ByVal path As String, h As Long) As Boolean
    Dim made As Long, rc As Long
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, path, 0, vbNullString, 0, KEY_ALL_ACCESS, 0, h, made)
    If rc = ERROR_SUCCESS Then
        OpenForWriting = True
    Else
        Refused path, rc
    End If
End Function

' The registry said no: say so once. On Windows XP a limited account can't write to HKEY_LOCAL_MACHINE.
Private Sub Refused(ByVal path As String, ByVal code As Long)
    If mBroken Then Exit Sub
    mBroken = True
    If code = ERROR_ACCESS_DENIED Then
        MsgBox "Palms LAN Center is unable to save its settings in the registry (HKEY_LOCAL_MACHINE\" & path & ")." & vbCrLf & vbCrLf & _
               "Log on to Windows as an administrator. Until then, settings and installed games are not remembered.", vbExclamation, "Palms LAN Center"
    Else
        MsgBox "Palms LAN Center is unable to save a setting in the registry (HKEY_LOCAL_MACHINE\" & path & ")." & vbCrLf & vbCrLf & _
               "Windows error " & code & ".", vbExclamation, "Palms LAN Center"
    End If
End Sub

' ================================================================ the old .ini files

' Reads palms.ini and gamepaks.ini (from older versions) into the registry, once.
' Afterwards they are renamed .old, so they are never read again. If the registry refused
' (a limited account on XP), they stay as they are, to be read in when an administrator runs it.
Public Sub ImportOldFiles()
    Dim f As String
    f = AppDir() & "palms.ini"
    If FileHere(f) Then
        ImportSection f, "client", REG_LAN
        If Not mBroken Then RenameOld f
    End If
    f = AppDir() & "gamepaks.ini"
    If FileHere(f) Then
        ImportGames f
        If Not mBroken Then RenameOld f
    End If
End Sub

' gamepaks.ini: [installed] list=, [quake], [halflifepack/counterstrike16]...
Private Sub ImportGames(ByVal f As String)
    Dim buf As String, n As Long, sects() As String, i As Long, s As String
    buf = String$(65536, vbNullChar)
    n = GetPrivateProfileSectionNames(buf, Len(buf), f)
    If n <= 0 Then Exit Sub
    sects = Split(Left$(buf, n), vbNullChar)
    For i = 0 To UBound(sects)
        s = Trim$(sects(i))
        If Len(s) > 0 Then
            If LCase$(s) = "installed" Then
                RegPut REG_GAMES, "Installed", ProfileValue(f, s, "list")
            Else
                ImportSection f, s, GamesKey(s)
            End If
        End If
    Next
    ' Add/Remove Programs for each of them
    sects = Split(RegGet(REG_GAMES, "Installed", ""), ",")
    For i = 0 To UBound(sects)
        s = LCase$(Trim$(sects(i)))
        If Len(s) > 0 Then WriteUninstallEntry s
    Next
End Sub

' Every key = value in one section, copied into a registry key
Private Sub ImportSection(ByVal f As String, ByVal sect As String, ByVal path As String)
    Dim buf As String, n As Long, lines() As String, i As Long, p As Long
    buf = String$(32767, vbNullChar)
    n = GetPrivateProfileSection(sect, buf, Len(buf), f)
    If n <= 0 Then Exit Sub
    lines = Split(Left$(buf, n), vbNullChar)
    For i = 0 To UBound(lines)
        p = InStr(lines(i), "=")
        If p > 1 Then RegPut path, Trim$(Left$(lines(i), p - 1)), Mid$(lines(i), p + 1)
    Next
End Sub

Private Function ProfileValue(ByVal f As String, ByVal sect As String, ByVal key As String) As String
    Dim buf As String, n As Long, lines() As String, i As Long, p As Long
    buf = String$(32767, vbNullChar)
    n = GetPrivateProfileSection(sect, buf, Len(buf), f)
    If n <= 0 Then Exit Function
    lines = Split(Left$(buf, n), vbNullChar)
    For i = 0 To UBound(lines)
        p = InStr(lines(i), "=")
        If p > 1 Then
            If LCase$(Trim$(Left$(lines(i), p - 1))) = LCase$(key) Then
                ProfileValue = Mid$(lines(i), p + 1)
                Exit Function
            End If
        End If
    Next
End Function

Private Sub RenameOld(ByVal f As String)
    On Error Resume Next
    If FileHere(f & ".old") Then Kill f & ".old"
    Name f As f & ".old"
End Sub

Public Function FileHere(ByVal f As String) As Boolean
    On Error Resume Next
    FileHere = ((GetAttr(f) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileHere = False
End Function

' ================================================================ the installed games

' "quake" is Games\quake, "halflifepack/counterstrike16" is Games\halflifepack\counterstrike16
Public Function GamesKey(ByVal sect As String) As String
    GamesKey = REG_GAMES & "\" & Replace(sect, "/", "\")
End Function

' Add/Remove Programs: the cartridge's name, and PalmsLAN.exe /uninstall <cartridge> to remove it
Public Sub WriteUninstallEntry(ByVal id As String)
    Dim k As String, g As String, nm As String, place As String, ico As String, games() As String, exe As String, kb As Double
    g = GamesKey(id)
    nm = RegGet(g, "name", "")
    place = RegGet(g, "folder", "")
    If Len(nm) = 0 Or Len(place) = 0 Then Exit Sub
    k = REG_UNINSTALL & "\PalmsArcade_" & id
    RegPut k, "DisplayName", nm
    RegPut k, "DisplayVersion", VerText(RegGet(g, "version", ""), RegGet(g, "revision", ""))
    RegPut k, "Publisher", "Palms Arcade"
    RegPut k, "Comments", "Installed by Palms LAN Center"
    RegPut k, "InstallLocation", place
    RegPut k, "UninstallString", """" & AppDir() & "PalmsLAN.exe"" /uninstall " & id
    ' the icon: the first game's program
    exe = RegGet(g, "exe", "")
    games = Split(RegGet(g, "games", ""), ",")
    If UBound(games) >= 0 Then
        If Len(Trim$(games(0))) > 0 Then exe = RegGet(GamesKey(id & "/" & Trim$(games(0))), "exe", exe)
    End If
    If Len(exe) > 0 Then
        ico = place & "\" & exe
        RegPut k, "DisplayIcon", ico
    End If
    RegPutNumber k, "NoModify", 1
    RegPutNumber k, "NoRepair", 1
    kb = Val(RegGet(g, "size", "0")) / 1024
    If kb > 0 And kb < 2147483647# Then RegPutNumber k, "EstimatedSize", CLng(kb)
End Sub

Public Sub DropUninstallEntry(ByVal id As String)
    RegDropKey REG_UNINSTALL & "\PalmsArcade_" & id
End Sub
