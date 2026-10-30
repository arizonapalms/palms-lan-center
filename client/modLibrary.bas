Attribute VB_Name = "modLibrary"
' Palms LAN Center - The Game Library: the cartridges on offer, the games installed here, and installing and removing them.
Option Explicit

' ---- Windows calls
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFileForChild Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const SW_HIDE As Integer = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000
Private Const WAIT_TIMEOUT As Long = 258
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE As Long = -1
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const CP_UTF8 As Long = 65001
Private Const KEY_READ As Long = &H20019

' ---- where things go
Public Const GAMES_DEFAULT As String = "C:\LANPartyGames"
Private Const SPARE_BYTES As Double = 52428800#    ' leave at least 50 MB free on the drive
Private Const NOTES_NAME As String = "Palms notes.txt"
Private Const MENU_NAME As String = "Palms Arcade Games"   ' the Start menu folder for the shortcuts

' ---- a game on the pak server
Public Type GamePak
    Id As String          ' short name, like quake
    Name As String        ' Quake
    Ver As String         ' v1.09
    Rev As String         ' the cartridge's revision: 2 for the second cartridge of the same game version
    Players As String     ' 2-16
    Folder As String      ' the folder it installs into, like Quake
    Exe As String         ' main program, inside that folder
    Args As String
    CfgExe As String      ' settings program, or ""
    CfgArgs As String
    Reg As String         ' setup.reg, or ""
    Notes As String       ' notes.txt, or ""
    Path As String        ' the pak's folder on the pak server, like paks/quake-v1.09/
    Archive As String     ' game.7z
    Pieces As Long
    ArcBytes As Double    ' all the pieces together: the download
    Size As Double        ' the game once it's installed
    Files As Long
    Icon As String        ' icons/quake-0ea49e6702.bmp on the pak server, or ""
    Shot As String        ' its screenshot in the icons folder, like quake-0ea49e6702-shot.jpg, or ""
    Made As String        ' when the pak was made: a different one means an update
End Type

' ---- a game installed on this PC (a key under HKLM\Software\Palms Arcade\Games)
Public Type GameHere
    Id As String
    Name As String
    Ver As String
    Rev As String
    Made As String
    Place As String       ' the folder it's in, like C:\LANPartyGames\Quake
    Exe As String
    Args As String
    CfgExe As String
    CfgArgs As String
    Size As Double
    Notes As String       ' the extra steps file in the game's folder, or ""
    NotesSeen As Boolean  ' the player has been offered them
    RegKeys As String     ' registry keys setup.reg made, separated by ;
    Icon As String        ' its icon in the icons folder, or ""
    Installed As String   ' when, like 2026-10-22 18:40
    Games As String       ' the games in it, like counterstrike16,opposingforce ("" for a single-game cartridge from an early version)
    Programs As String    ' its programs, let through the firewall, separated by ;
    Shortcuts As String   ' its Start menu shortcuts, separated by ;
End Type

' ---- one game in a pak. Most paks are one game; a pack (like a Half-Life folder with
' Counter-Strike and Opposing Force in it) is several, all installed together.
Public Type GameInfo
    Key As String         ' pak id/game id, like halflifepack/counterstrike16 (or quake/quake)
    PakId As String
    Id As String
    Name As String
    Players As String
    Exe As String         ' its program, inside the pak's folder
    Args As String        ' what goes after it, like -game cstrike
    StartIn As String     ' the folder inside the pak's folder it starts in, or "" for the program's own
    CfgExe As String
    CfgArgs As String
    Icon As String        ' its icon file in the icons folder, like halflifepack-counterstrike16-d020d7a2ed.bmp
End Type

Public gPak() As GamePak
Public gPakCount As Long
Public gPakBase As String      ' the pak server's address, like http://10.0.200.73/
Public gPakNote As String      ' from the Palms server: why the list is empty or out of date
Public gPakHeard As Boolean    ' the Palms server has sent a list

Public gHere() As GameHere
Public gHereCount As Long

Public gGame() As GameInfo     ' the games on the pak server
Public gGameCount As Long
Private mHereGame() As GameInfo    ' the games installed here
Private mHereGameCount As Long

' Every game the page shows: the pak server's, then installed ones that aren't on it any more
Public gShowId() As String
Public gShowCount As Long

' ---- what a game is up to
Public Const LS_NONE As Integer = 0       ' not installed
Public Const LS_READY As Integer = 1      ' installed, the same pak the pak server has
Public Const LS_UPDATE As Integer = 2     ' installed, and the pak server has a newer pak
Public Const LS_OLD As Integer = 3        ' installed, but no longer on the pak server
Public Const LS_WAIT As Integer = 4       ' waiting its turn
Public Const LS_BUSY As Integer = 5       ' being installed or uninstalled right now

' ---- the job in hand
Public gJobId As String          ' the pak being installed or uninstalled, "" when there isn't one
Public gJobPct As Long           ' how far along, 0 to 100, or -1 when we can't tell
Public gJobText As String        ' what's happening, for the card
Public gJobShort As String       ' the same, short, for the list
Public gLibDirty As Boolean      ' a game was installed or removed: the list needs sorting again
Public gIconsNew As Boolean      ' new icons have arrived
Public gFailId As String         ' the pak whose last try went wrong...
Public gFailText As String       ' ...and what went wrong

Private Const ST_IDLE As Integer = 0
Private Const ST_SMALL As Integer = 1     ' downloading setup.reg and notes.txt
Private Const ST_PIECES As Integer = 2    ' downloading the game
Private Const ST_CLEAR As Integer = 3     ' deleting the old copy
Private Const ST_UNPACK As Integer = 4    ' 7-Zip unpacking
Private Const ST_REG As Integer = 5       ' merging setup.reg
Private Const ST_WALL As Integer = 6      ' the firewall exception
Private Const ST_REMOVE As Integer = 7    ' uninstalling: deleting the folder
Private Const ST_UNWALL As Integer = 8    ' uninstalling: the firewall exception
Private Const ST_UNREG As Integer = 9     ' uninstalling: the registry keys

Private mStep As Integer
Private mRemoving As Boolean     ' uninstalling rather than installing
Private mProc As Long            ' the program doing the work (a process handle), 0 for none
Private mPak As GamePak          ' the game being installed, as it was when we started
Private mTarget As String        ' its folder, like C:\LANPartyGames\Quake
Private mDl As String            ' where its pieces download to
Private mFreeAtStart As Double   ' free space before unpacking: what's gone since is how far along
Private mSpeed As Double         ' download speed, bytes a second, smoothed
Private mLastBytes As Double
Private mLastTick As Long
Private mRegKeys As String       ' the game's own registry keys, to take away when it's uninstalled
Private mRegCheck As String      ' every key setup.reg puts settings in, to check they went in
Private mRegWarn As String       ' they didn't all go in: what to tell the player
Private mQueue() As String       ' "i" to install or "u" to uninstall, then the pak's id: iquake, uhalflifepack
Private mQueueCount As Long
Private mIconProc As Long        ' palmsget fetching icons, 0 for none
Private mIcoTried As String      ' .ico files already asked for, so a pak server without them isn't asked every time
Private mPakGames() As GameInfo  ' the games in the pak being installed, as they were when we started
Private mPakGameCount As Long
Private mWall() As String        ' programs to let through (or take off) the firewall, one at a time
Private mWallCount As Long
Private mWallIdx As Long

' ================================================================ places

Public Function GamesRoot() As String
    Dim s As String
    s = Trim$(SettingGet("games_folder", GAMES_DEFAULT))
    If Len(s) = 0 Then s = GAMES_DEFAULT
    If Len(s) > 3 And Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    GamesRoot = s
End Function

' "C:\LANPartyGames" + "Quake", and "C:\" + "Quake" without doubling the \
Private Function JoinPath(ByVal a As String, ByVal b As String) As String
    If Right$(a, 1) = "\" Then
        JoinPath = a & b
    Else
        JoinPath = a & "\" & b
    End If
End Function

' "C:\LANPartyGames\Quake" becomes "C:\LANPartyGames", "C:\Quake" becomes "C:"
Private Function ParentOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then ParentOf = Left$(p, i - 1)
End Function

Private Function TempDir() As String
    TempDir = AppDir() & "temp"
    MakeDirs TempDir
End Function

Private Function IconDir() As String
    IconDir = AppDir() & "icons"
End Function

Private Function WinDir() As String
    Dim buf As String, n As Long
    buf = String$(260, vbNullChar)
    n = GetWindowsDirectory(buf, 260)
    WinDir = Left$(buf, n)
    If Len(WinDir) = 0 Then WinDir = "C:\WINDOWS"
End Function

' The two helper programs, or "" when they're both here
Public Function ToolsMissing() As String
    Dim a As Boolean, b As Boolean
    a = FileThere(AppDir() & "palmsget.exe")
    b = FileThere(AppDir() & "7za.exe")
    If Not a And Not b Then
        ToolsMissing = "palmsget.exe and 7za.exe"
    ElseIf Not a Then
        ToolsMissing = "palmsget.exe"
    ElseIf Not b Then
        ToolsMissing = "7za.exe"
    End If
End Function

Public Function FileThere(ByVal path As String) As Boolean
    On Error GoTo Nope
    FileThere = ((GetAttr(path) And vbDirectory) = 0)
Nope:
End Function

Public Function FolderThere(ByVal path As String) As Boolean
    On Error GoTo Nope
    FolderThere = ((GetAttr(path) And vbDirectory) <> 0)
Nope:
End Function

' Makes a folder, and the folders it's in
Public Sub MakeDirs(ByVal path As String)
    Dim p As Long
    On Error Resume Next
    If Len(path) < 3 Then Exit Sub
    If FolderThere(path) Then Exit Sub
    p = InStrRev(path, "\")
    If p > 3 Then MakeDirs Left$(path, p - 1)
    MkDir path
End Sub

' Free space on the drive a folder is on, in bytes. -1 if Windows can't say.
Public Function FreeBytes(ByVal path As String) As Double
    Dim a As Currency, b As Currency, c As Currency
    On Error GoTo Nope
    FreeBytes = -1
    If Not (path Like "[A-Za-z]:*") Then Exit Function
    If GetDiskFreeSpaceEx(Left$(path, 2) & "\", a, b, c) = 0 Then Exit Function
    ' Currency counts in ten-thousandths, so the bytes come back divided by 10,000
    FreeBytes = CDbl(a) * 10000#
Nope:
End Function

' How big a file is right now, even while another program is still writing it
Private Function FileBytes(ByVal path As String) As Double
    Dim h As Long, lo As Long, hi As Long
    h = CreateFile(path, 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If h = INVALID_HANDLE Then h = CreateFile(path, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If h = INVALID_HANDLE Then Exit Function
    lo = GetFileSize(h, hi)
    CloseHandle h
    If lo = INVALID_HANDLE And hi = 0 Then Exit Function
    FileBytes = Unsigned(lo) + CDbl(hi) * 4294967296#
End Function

Private Function Unsigned(ByVal n As Long) As Double
    If n < 0 Then
        Unsigned = CDbl(n) + 4294967296#
    Else
        Unsigned = n
    End If
End Function

' "390 KB", "490 MB", "1.2 GB"
Public Function SizeText(ByVal n As Double) As String
    If n >= 1073741824# Then
        SizeText = Format$(n / 1073741824#, "0.0") & " GB"
    ElseIf n >= 1048576# Then
        SizeText = Format$(n / 1048576#, "0") & " MB"
    ElseIf n >= 1024# Then
        SizeText = Format$(n / 1024#, "0") & " KB"
    ElseIf n > 0 Then
        SizeText = "1 KB"
    Else
        SizeText = "0 KB"
    End If
End Function

Private Function Quoted(ByVal s As String) As String
    Quoted = """" & s & """"
End Function

' Folder names from the pak server: letters, numbers, - _ and . only, so nothing can
' point outside the games folder
Private Function SafeName(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Or Len(s) > 64 Then Exit Function
    If s = "." Or s = ".." Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not (ch Like "[A-Za-z0-9._-]") Then Exit Function
    Next
    SafeName = True
End Function

' A program inside a game's folder, like quake.exe or System\UnrealTournament.exe
Private Function SafeInside(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    If InStr(s, "..") > 0 Or InStr(s, ":") > 0 Or Left$(s, 1) = "\" Then Exit Function
    SafeInside = True
End Function

' Never a drive, the Windows folder, Program Files, the LAN Center's folder or the games folder itself
Private Function SafeToDelete(ByVal p As String) As Boolean
    Dim low As String
    low = LCase$(p)
    If Len(p) < 4 Or Mid$(p, 2, 2) <> ":\" Or InStr(p, "..") > 0 Or Right$(p, 1) = "\" Then Exit Function
    If low = LCase$(WinDir()) Or low = LCase$(Environ$("ProgramFiles")) Then Exit Function
    If low & "\" = LCase$(AppDir()) Or low = LCase$(GamesRoot()) Then Exit Function
    SafeToDelete = True
End Function

' ================================================================ small files

Private Function ReadText(ByVal path As String) As String
    Dim fnum As Integer, n As Long, s As String
    On Error GoTo Nope
    fnum = FreeFile
    Open path For Binary Access Read Shared As #fnum
    n = LOF(fnum)
    If n > 0 Then
        s = String$(n, vbNullChar)
        Get #fnum, , s
    End If
    Close #fnum
    ReadText = s
    Exit Function
Nope:
    On Error Resume Next
    Close #fnum
End Function

' The end of a log file, for saying what went wrong
Private Function ReadTail(ByVal path As String) As String
    Dim fnum As Integer, n As Long, s As String, start As Long
    On Error GoTo Nope
    fnum = FreeFile
    Open path For Binary Access Read Shared As #fnum
    n = LOF(fnum)
    start = n - 4096
    If start < 0 Then start = 0
    If n > start Then
        s = String$(n - start, vbNullChar)
        Get #fnum, start + 1, s
    End If
    Close #fnum
    ReadTail = s
    Exit Function
Nope:
    On Error Resume Next
    Close #fnum
End Function

Private Sub WriteText(ByVal path As String, ByVal s As String)
    Dim fnum As Integer
    On Error GoTo Nope
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, s;
    Close #fnum
    Exit Sub
Nope:
    On Error Resume Next
    Close #fnum
End Sub

Private Sub KillQuiet(ByVal pattern As String)
    On Error Resume Next
    If Right$(pattern, 1) = "\" Then Exit Sub
    Kill pattern
End Sub

Private Sub RmDirQuiet(ByVal path As String)
    On Error Resume Next
    RmDir path
End Sub

' The last line of a log that isn't blank
Private Function LastLine(ByVal text As String) As String
    Dim lines() As String, i As Long
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For i = UBound(lines) To 0 Step -1
        If Len(Trim$(lines(i))) > 0 Then
            LastLine = Trim$(lines(i))
            Exit Function
        End If
    Next
End Function

' ================================================================ what's installed, in the registry
' Under HKEY_LOCAL_MACHINE\Software\Palms Arcade\Games:
'   Installed                        quake,halflifepack
'   \halflifepack                    the pak: name, version, made, folder, notes, regkeys,
'                                    programs (for the firewall), shortcuts, games=...
'   \halflifepack\counterstrike16    one of its games: name, players, exe, args, start_in...
' A pak installed by an early version has no games value: it's one game, its program in \quake.
' "installed" and "halflifepack/counterstrike16" are what gamepaks.ini called its sections.

Private Function PIni(ByVal sect As String, ByVal key As String, ByVal def As String) As String
    If sect = "installed" Then
        PIni = RegGet(REG_GAMES, "Installed", def)
    Else
        PIni = RegGet(GamesKey(sect), key, def)
    End If
End Function

Private Sub PIniSet(ByVal sect As String, ByVal key As String, ByVal value As String)
    If sect = "installed" Then
        RegPut REG_GAMES, "Installed", value
    Else
        RegPut GamesKey(sect), key, value
    End If
End Sub

' Forgets a pak, or one of its games, and everything under it
Private Sub PIniDrop(ByVal sect As String)
    RegDropKey GamesKey(sect)
End Sub

Public Sub LoadHere()
    Dim ids() As String, i As Long, id As String, games() As String, k As Long, sect As String, gid As String
    ids = Split(PIni("installed", "list", ""), ",")
    ReDim gHere(0 To UBound(ids) + 1)
    ReDim mHereGame(0 To 7)
    gHereCount = 0
    mHereGameCount = 0
    For i = 0 To UBound(ids)
        id = LCase$(Trim$(ids(i)))
        If Len(id) > 0 And HereIndex(id) < 0 Then
            If Len(PIni(id, "folder", "")) > 0 Then
                With gHere(gHereCount)
                    .Id = id
                    .Name = PIni(id, "name", id)
                    .Ver = PIni(id, "version", "")
                    .Rev = PIni(id, "revision", "")
                    .Made = PIni(id, "made", "")
                    .Place = PIni(id, "folder", "")
                    .Exe = PIni(id, "exe", "")
                    .Args = PIni(id, "args", "")
                    .CfgExe = PIni(id, "config_exe", "")
                    .CfgArgs = PIni(id, "config_args", "")
                    .Size = Val(PIni(id, "size", "0"))
                    .Notes = PIni(id, "notes", "")
                    .NotesSeen = (PIni(id, "notes_seen", "0") = "1")
                    .RegKeys = PIni(id, "regkeys", "")
                    .Icon = PIni(id, "icon", "")
                    .Installed = PIni(id, "installed", "")
                    .Games = PIni(id, "games", "")
                    .Programs = PIni(id, "programs", "")
                    .Shortcuts = PIni(id, "shortcuts", "")
                    If Len(.Programs) = 0 Then .Programs = .Exe
                End With
                gHereCount = gHereCount + 1
                With gHere(gHereCount - 1)
                    If Len(.Games) = 0 Then
                        ' installed by an early version: one game, the pak's own program
                        If Len(.Exe) > 0 Then AddHereGame id, id, .Name, PIni(id, "players", ""), .Exe, .Args, "", .CfgExe, .CfgArgs, .Icon
                    Else
                        games = Split(.Games, ",")
                        For k = 0 To UBound(games)
                            gid = Trim$(games(k))
                            sect = id & "/" & gid
                            If Len(gid) > 0 And Len(PIni(sect, "exe", "")) > 0 Then
                                AddHereGame id, gid, PIni(sect, "name", gid), PIni(sect, "players", ""), PIni(sect, "exe", ""), _
                                            PIni(sect, "args", ""), PIni(sect, "start_in", ""), PIni(sect, "config_exe", ""), _
                                            PIni(sect, "config_args", ""), PIni(sect, "icon", "")
                            End If
                        Next
                    End If
                End With
            End If
        End If
    Next
End Sub

Private Sub AddHereGame(ByVal pak As String, ByVal gid As String, ByVal nm As String, ByVal players As String, ByVal exe As String, _
                        ByVal args As String, ByVal startIn As String, ByVal cfg As String, ByVal cfgArgs As String, ByVal icon As String)
    If mHereGameCount = 0 Then ReDim mHereGame(0 To 7)
    If mHereGameCount > UBound(mHereGame) Then ReDim Preserve mHereGame(0 To mHereGameCount + 8)
    With mHereGame(mHereGameCount)
        .PakId = pak
        .Id = gid
        .Key = pak & "/" & gid
        .Name = nm
        .Players = players
        .Exe = exe
        .Args = args
        .StartIn = startIn
        .CfgExe = cfg
        .CfgArgs = cfgArgs
        .Icon = icon
    End With
    mHereGameCount = mHereGameCount + 1
End Sub

Private Sub WriteList()
    Dim i As Long, s As String
    For i = 0 To gHereCount - 1
        If i > 0 Then s = s & ","
        s = s & gHere(i).Id
    Next
    PIniSet "installed", "list", s
End Sub

' Records a pak as installed, with its games (from mPakGames, the ones it was installed with)
Private Sub PutHere(h As GameHere)
    Dim i As Long, k As Long, sect As String
    Dim old() As String
    i = HereIndex(h.Id)
    If i < 0 Then
        If gHereCount > UBound(gHere) Then ReDim Preserve gHere(0 To gHereCount + 4)
        i = gHereCount
        gHereCount = gHereCount + 1
    Else
        ' the old copy's games
        old = Split(gHere(i).Games, ",")
        For k = 0 To UBound(old)
            If Len(Trim$(old(k))) > 0 Then PIniDrop h.Id & "/" & Trim$(old(k))
        Next
    End If
    gHere(i) = h
    gLibDirty = True
    PIniSet h.Id, "name", h.Name
    PIniSet h.Id, "version", h.Ver
    PIniSet h.Id, "revision", h.Rev
    PIniSet h.Id, "made", h.Made
    PIniSet h.Id, "folder", h.Place
    PIniSet h.Id, "exe", h.Exe
    PIniSet h.Id, "args", h.Args
    PIniSet h.Id, "config_exe", h.CfgExe
    PIniSet h.Id, "config_args", h.CfgArgs
    PIniSet h.Id, "size", Format$(h.Size, "0")
    PIniSet h.Id, "notes", h.Notes
    If h.NotesSeen Then
        PIniSet h.Id, "notes_seen", "1"
    Else
        PIniSet h.Id, "notes_seen", "0"
    End If
    PIniSet h.Id, "regkeys", h.RegKeys
    PIniSet h.Id, "icon", h.Icon
    PIniSet h.Id, "installed", h.Installed
    PIniSet h.Id, "games", h.Games
    PIniSet h.Id, "programs", h.Programs
    PIniSet h.Id, "shortcuts", h.Shortcuts
    ' its games
    DropHereGames h.Id
    For k = 0 To mPakGameCount - 1
        With mPakGames(k)
            sect = h.Id & "/" & .Id
            PIniSet sect, "name", .Name
            PIniSet sect, "players", .Players
            PIniSet sect, "exe", .Exe
            PIniSet sect, "args", .Args
            PIniSet sect, "start_in", .StartIn
            PIniSet sect, "config_exe", .CfgExe
            PIniSet sect, "config_args", .CfgArgs
            PIniSet sect, "icon", .Icon
            AddHereGame h.Id, .Id, .Name, .Players, .Exe, .Args, .StartIn, .CfgExe, .CfgArgs, .Icon
        End With
    Next
    WriteList
    ' and in Add/Remove Programs
    WriteUninstallEntry h.Id
End Sub

Private Sub DropHereGames(ByVal pak As String)
    Dim i As Long, n As Long
    For i = 0 To mHereGameCount - 1
        If mHereGame(i).PakId <> pak Then
            If n <> i Then mHereGame(n) = mHereGame(i)
            n = n + 1
        End If
    Next
    mHereGameCount = n
End Sub

Private Sub DropHere(ByVal id As String)
    Dim i As Long, k As Long, games() As String
    i = HereIndex(id)
    If i < 0 Then Exit Sub
    games = Split(gHere(i).Games, ",")
    For k = 0 To UBound(games)
        If Len(Trim$(games(k))) > 0 Then PIniDrop id & "/" & Trim$(games(k))
    Next
    For k = i To gHereCount - 2
        gHere(k) = gHere(k + 1)
    Next
    gHereCount = gHereCount - 1
    DropHereGames id
    gLibDirty = True
    PIniDrop id
    DropUninstallEntry id
    WriteList
End Sub

Public Function HereIndex(ByVal id As String) As Long
    Dim i As Long
    HereIndex = -1
    For i = 0 To gHereCount - 1
        If gHere(i).Id = id Then
            HereIndex = i
            Exit Function
        End If
    Next
End Function

Public Function CatIndex(ByVal id As String) As Long
    Dim i As Long
    CatIndex = -1
    For i = 0 To gPakCount - 1
        If gPak(i).Id = id Then
            CatIndex = i
            Exit Function
        End If
    Next
End Function

' A game on the pak server, by its key (pak id/game id)
Private Function GameIndex(ByVal key As String) As Long
    Dim i As Long
    GameIndex = -1
    For i = 0 To gGameCount - 1
        If gGame(i).Key = key Then
            GameIndex = i
            Exit Function
        End If
    Next
End Function

' A game installed here, by its key
Private Function HereGameIndex(ByVal key As String) As Long
    Dim i As Long
    HereGameIndex = -1
    For i = 0 To mHereGameCount - 1
        If mHereGame(i).Key = key Then
            HereGameIndex = i
            Exit Function
        End If
    Next
End Function

' The pak a game is in: halflifepack/counterstrike16 is in halflifepack. A pak id stays itself.
Public Function PakOf(ByVal key As String) As String
    Dim p As Long
    p = InStr(key, "/")
    If p > 0 Then
        PakOf = Left$(key, p - 1)
    Else
        PakOf = key
    End If
End Function

' ================================================================ the list from the pak server

Private Function F2(f() As String, ByVal i As Integer) As String
    If i <= UBound(f) Then F2 = Trim$(f(i))
End Function

' PAKS|how many|pak server address|note
Public Sub PaksBegin(f() As String)
    Dim n As Long
    n = Val(F2(f, 1))
    If n < 0 Then n = 0
    ReDim gPak(0 To n)
    gPakCount = 0
    ReDim gGame(0 To n + 8)
    gGameCount = 0
    gPakBase = F2(f, 2)
    If Len(gPakBase) > 0 And Right$(gPakBase, 1) <> "/" Then gPakBase = gPakBase & "/"
    gPakNote = F2(f, 3)
End Sub

' PAK|id|name|version|players|folder|exe|args|config exe|config args|reg|notes|
'     pak folder|archive|pieces|archive bytes|installed size|files|icon|made
' A pack's exe is blank: its games come in GAME lines after it.
Public Sub PaksAdd(f() As String)
    Dim g As GamePak
    g.Id = LCase$(F2(f, 1))
    g.Name = F2(f, 2)
    g.Ver = F2(f, 3)
    g.Players = F2(f, 4)
    g.Folder = F2(f, 5)
    g.Exe = F2(f, 6)
    g.Args = F2(f, 7)
    g.CfgExe = F2(f, 8)
    g.CfgArgs = F2(f, 9)
    g.Reg = F2(f, 10)
    g.Notes = F2(f, 11)
    g.Path = F2(f, 12)
    g.Archive = F2(f, 13)
    g.Pieces = Val(F2(f, 14))
    g.ArcBytes = Val(F2(f, 15))
    g.Size = Val(F2(f, 16))
    g.Files = Val(F2(f, 17))
    g.Icon = F2(f, 18)
    g.Made = F2(f, 19)
    g.Rev = F2(f, 20)                ' blank from older Cartridge Servers
    If Not (g.Rev Like "#" Or g.Rev Like "##" Or g.Rev Like "###") Then g.Rev = ""
    g.Shot = IconFileOf(F2(f, 21))                ' blank from older Cartridge Servers
    If Len(g.Name) = 0 Then g.Name = g.Id
    If Len(g.Archive) = 0 Then g.Archive = "game.7z"
    If Len(g.Path) > 0 And Right$(g.Path, 1) <> "/" Then g.Path = g.Path & "/"
    ' anything odd and the pak is left out, rather than installed somewhere strange
    If Len(g.Id) = 0 Or g.Pieces < 1 Or Not SafeName(g.Folder) Then Exit Sub
    If Len(g.Exe) > 0 And Not SafeInside(g.Exe) Then Exit Sub
    If Len(g.CfgExe) > 0 And Not SafeInside(g.CfgExe) Then g.CfgExe = ""
    If Len(g.Reg) > 0 And Not SafeName(g.Reg) Then g.Reg = ""
    If Len(g.Notes) > 0 And Not SafeName(g.Notes) Then g.Notes = ""
    If Not SafeName(g.Archive) Then Exit Sub
    If CatIndex(g.Id) >= 0 Then Exit Sub
    If gPakCount > UBound(gPak) Then ReDim Preserve gPak(0 To gPakCount + 8)
    gPak(gPakCount) = g
    gPakCount = gPakCount + 1
End Sub

' GAME|pak id|game id|name|players|exe|args|start in|config exe|config args|icon
' One game in a pack (Counter-Strike in a Half-Life pack, say)
Public Sub PaksGame(f() As String)
    Dim g As GameInfo
    g.PakId = LCase$(F2(f, 1))
    g.Id = LCase$(F2(f, 2))
    g.Name = F2(f, 3)
    g.Players = F2(f, 4)
    g.Exe = F2(f, 5)
    g.Args = F2(f, 6)
    g.StartIn = F2(f, 7)
    g.CfgExe = F2(f, 8)
    g.CfgArgs = F2(f, 9)
    g.Icon = IconFileOf(F2(f, 10))
    If CatIndex(g.PakId) < 0 Then Exit Sub
    If Not SafeName(g.Id) Or Not SafeInside(g.Exe) Then Exit Sub
    If Len(g.StartIn) > 0 And Not SafeInside(g.StartIn) Then g.StartIn = ""
    If Len(g.CfgExe) > 0 And Not SafeInside(g.CfgExe) Then g.CfgExe = ""
    If Len(g.Name) = 0 Then g.Name = g.Id
    g.Key = g.PakId & "/" & g.Id
    If GameIndex(g.Key) >= 0 Then Exit Sub
    AddGame g
End Sub

Private Sub AddGame(g As GameInfo)
    If gGameCount = 0 Then ReDim gGame(0 To 7)
    If gGameCount > UBound(gGame) Then ReDim Preserve gGame(0 To gGameCount + 8)
    gGame(gGameCount) = g
    gGameCount = gGameCount + 1
End Sub

Public Sub PaksEnd()
    Dim i As Long, k As Long, has As Boolean, g As GameInfo
    ' a pak without GAME lines is one game: the pak's own program
    For i = 0 To gPakCount - 1
        has = False
        For k = 0 To gGameCount - 1
            If gGame(k).PakId = gPak(i).Id Then has = True
        Next
        If Not has And Len(gPak(i).Exe) > 0 Then
            g.PakId = gPak(i).Id
            g.Id = gPak(i).Id
            g.Key = g.PakId & "/" & g.Id
            g.Name = gPak(i).Name
            g.Players = gPak(i).Players
            g.Exe = gPak(i).Exe
            g.Args = gPak(i).Args
            g.StartIn = ""
            g.CfgExe = gPak(i).CfgExe
            g.CfgArgs = gPak(i).CfgArgs
            g.Icon = IconFileOf(gPak(i).Icon)
            AddGame g
        End If
    Next
    gPakHeard = True
    LibRebuild
End Sub

' Every game to show: the pak server's, then installed ones it doesn't have any more
Public Sub LibRebuild()
    Dim i As Long
    ReDim gShowId(0 To gGameCount + mHereGameCount)
    gShowCount = 0
    For i = 0 To gGameCount - 1
        gShowId(gShowCount) = gGame(i).Key
        gShowCount = gShowCount + 1
    Next
    For i = 0 To mHereGameCount - 1
        If GameIndex(mHereGame(i).Key) < 0 Then
            gShowId(gShowCount) = mHereGame(i).Key
            gShowCount = gShowCount + 1
        End If
    Next
    gLibDirty = True
End Sub

' ================================================================ about one game, for the page
' Everything here takes a game's key (halflifepack/counterstrike16). Things that belong to
' the whole pak, like its folder, version and size, come from the pak it's in.

Public Function LibName(ByVal key As String) As String
    Dim g As Long
    g = GameIndex(key)
    If g >= 0 Then
        LibName = gGame(g).Name
        Exit Function
    End If
    g = HereGameIndex(key)
    If g >= 0 Then
        LibName = mHereGame(g).Name
    Else
        LibName = LibPakName(key)
    End If
End Function

' The pak's name: for a pack, like Half-Life Pack
Public Function LibPakName(ByVal key As String) As String
    Dim c As Long, h As Long, pak As String
    pak = PakOf(key)
    c = CatIndex(pak)
    If c >= 0 Then
        LibPakName = gPak(c).Name
        Exit Function
    End If
    h = HereIndex(pak)
    If h >= 0 Then
        LibPakName = gHere(h).Name
    Else
        LibPakName = pak
    End If
End Function

' The version on this PC if it's installed, otherwise the pak server's: 1.09 (Rev. 2)
Public Function LibVer(ByVal key As String) As String
    Dim c As Long, h As Long
    h = HereIndex(PakOf(key))
    If h >= 0 Then
        LibVer = VerText(gHere(h).Ver, gHere(h).Rev)
        Exit Function
    End If
    c = CatIndex(PakOf(key))
    If c >= 0 Then LibVer = VerText(gPak(c).Ver, gPak(c).Rev)
End Function

Public Function LibPlayers(ByVal key As String) As String
    Dim g As Long
    g = GameIndex(key)
    If g >= 0 Then
        LibPlayers = gGame(g).Players
        Exit Function
    End If
    g = HereGameIndex(key)
    If g >= 0 Then LibPlayers = mHereGame(g).Players
End Function

' How big the game (or the pack it's in) is once installed
Public Function LibSize(ByVal key As String) As Double
    Dim c As Long, h As Long
    c = CatIndex(PakOf(key))
    If c >= 0 Then
        LibSize = gPak(c).Size
        Exit Function
    End If
    h = HereIndex(PakOf(key))
    If h >= 0 Then LibSize = gHere(h).Size
End Function

' The game's icon file in the icons folder, or "" if it hasn't got one
Public Function LibIconFile(ByVal key As String) As String
    Dim g As Long
    g = GameIndex(key)
    If g >= 0 Then
        LibIconFile = gGame(g).Icon
        Exit Function
    End If
    g = HereGameIndex(key)
    If g >= 0 Then LibIconFile = mHereGame(g).Icon
End Function

' The screenshot of the cartridge a game is on, once it has been downloaded, or ""
Public Function LibShotPath(ByVal key As String) As String
    Dim c As Long
    c = CatIndex(PakOf(key))
    If c < 0 Then Exit Function
    If Len(gPak(c).Shot) = 0 Then Exit Function
    If FileThere(IconDir() & "\" & gPak(c).Shot) Then LibShotPath = IconDir() & "\" & gPak(c).Shot
End Function

' Every game on the Cartridge Server, sorted by name: key, a tab, then the name, one a line (vbLf between)
Public Function LibGameList() As String
    Dim i As Long, k As Long, order() As Long, t As Long, s As String
    If gGameCount = 0 Then Exit Function
    ReDim order(0 To gGameCount - 1)
    For i = 0 To gGameCount - 1
        order(i) = i
    Next
    ' few enough games that a simple sort will do
    For i = 1 To gGameCount - 1
        k = i
        Do While k > 0
            If StrComp(gGame(order(k - 1)).Name, gGame(order(k)).Name, vbTextCompare) <= 0 Then Exit Do
            t = order(k - 1)
            order(k - 1) = order(k)
            order(k) = t
            k = k - 1
        Loop
    Next
    For i = 0 To gGameCount - 1
        If Len(s) > 0 Then s = s & vbLf
        s = s & gGame(order(i)).Key & vbTab & gGame(order(i)).Name
    Next
    LibGameList = s
End Function

' A game the Cartridge Server has (its key, like halflifepack/counterstrike16)
Public Function LibOnServer(ByVal key As String) As Boolean
    LibOnServer = (GameIndex(key) >= 0)
End Function

Public Function LibIconPath(ByVal key As String) As String
    Dim f As String
    f = LibIconFile(key)
    If Len(f) > 0 Then LibIconPath = IconDir() & "\" & f
End Function

' icons/quake-0ea49e6702.bmp becomes quake-0ea49e6702.bmp
Private Function IconFileOf(ByVal url As String) As String
    Dim i As Long
    i = InStrRev(url, "/")
    IconFileOf = Mid$(url, i + 1)
    If Not SafeName(IconFileOf) Then IconFileOf = ""
End Function

' The pak is being installed or uninstalled, or waiting its turn
Private Function PakBusy(ByVal pak As String) As Boolean
    PakBusy = (pak = gJobId Or QueuePos(pak) >= 0)
End Function

Public Function LibState(ByVal key As String) As Integer
    Dim c As Long, h As Long, pak As String
    If Len(key) = 0 Then Exit Function
    pak = PakOf(key)
    If pak = gJobId Then
        LibState = LS_BUSY
        Exit Function
    End If
    If QueuePos(pak) >= 0 Then
        LibState = LS_WAIT
        Exit Function
    End If
    h = HereIndex(pak)
    If h < 0 Then Exit Function
    c = CatIndex(pak)
    If HereGameIndex(key) < 0 Then
        ' a game the installed copy of the pack doesn't have: the newer pak does
        If c >= 0 Then
            If gPak(c).Made <> gHere(h).Made Then LibState = LS_UPDATE
        End If
        Exit Function
    End If
    If c < 0 Then
        LibState = LS_OLD
    ElseIf gPak(c).Made <> gHere(h).Made Or gPak(c).Ver <> gHere(h).Ver Or gPak(c).Rev <> gHere(h).Rev Then
        LibState = LS_UPDATE
    Else
        LibState = LS_READY
    End If
End Function

Public Function LibInstalled(ByVal key As String) As Boolean
    LibInstalled = (HereGameIndex(key) >= 0)
End Function

' The pak (any of its games) is installed on this PC
Public Function LibPakInstalled(ByVal key As String) As Boolean
    LibPakInstalled = (HereIndex(PakOf(key)) >= 0)
End Function

' The last try at installing the pak this game is in went wrong (or needs a hand)
Public Function LibFailed(ByVal key As String) As Boolean
    LibFailed = (Len(gFailId) > 0 And PakOf(key) = gFailId)
End Function

' The pak server has the pak this game is in, so it can be installed
Public Function LibCanInstall(ByVal key As String) As Boolean
    LibCanInstall = (CatIndex(PakOf(key)) >= 0)
End Function

' How many games are in the pak this one is in (1 for most)
Public Function LibPackCount(ByVal key As String) As Long
    Dim i As Long, pak As String, n As Long
    pak = PakOf(key)
    For i = 0 To gShowCount - 1
        If PakOf(gShowId(i)) = pak Then n = n + 1
    Next
    LibPackCount = n
End Function

' The other games in the same pack, like "Half-Life, Opposing Force", or "" for a pak of one
Public Function LibPackOthers(ByVal key As String) As String
    Dim i As Long, pak As String, s As String
    pak = PakOf(key)
    For i = 0 To gShowCount - 1
        If PakOf(gShowId(i)) = pak And gShowId(i) <> key Then
            If Len(s) > 0 Then s = s & ", "
            s = s & LibName(gShowId(i))
        End If
    Next
    LibPackOthers = s
End Function

' What the list says about a game: words, their colour, an icon, and a progress bar (-1 for none)
Public Sub LibStatus(ByVal key As String, s As String, col As Long, ic As Integer, pct As Long)
    pct = -1
    Select Case LibState(key)
    Case LS_BUSY
        s = gJobShort
        col = C_BLUE
        If mRemoving Then
            ic = IC_ERASER
        Else
            ic = IC_DOWNLOAD
        End If
        pct = gJobPct
    Case LS_WAIT
        s = "Queued"
        col = C_GREY
        ic = IC_CLOCK
    Case LS_READY, LS_OLD
        If LibFailed(key) Then
            s = "Installed (errors)"
            col = C_ORANGE
            ic = IC_WARN
        Else
            s = "Installed"
            col = C_GREEN
            ic = IC_OK
        End If
    Case LS_UPDATE
        s = "Update available"
        col = C_ORANGE
        ic = IC_REFRESH
    Case Else
        If LibFailed(key) Then
            s = "Installation failed"
            col = C_RED
            ic = IC_WARN
        Else
            s = "Not installed"
            col = C_GREY
            ic = IC_BOX
        End If
    End Select
End Sub

' Where a game is, or where it will go (the pak's folder)
Public Function LibPlace(ByVal key As String) As String
    Dim c As Long, h As Long
    h = HereIndex(PakOf(key))
    If h >= 0 Then
        LibPlace = gHere(h).Place
        Exit Function
    End If
    c = CatIndex(PakOf(key))
    If c >= 0 Then LibPlace = JoinPath(GamesRoot(), gPak(c).Folder)
End Function

Public Function LibHasSettings(ByVal key As String) As Boolean
    Dim g As Long
    g = HereGameIndex(key)
    If g >= 0 Then LibHasSettings = (Len(mHereGame(g).CfgExe) > 0)
End Function

Public Function LibHasNotes(ByVal key As String) As Boolean
    Dim h As Long
    h = HereIndex(PakOf(key))
    If h >= 0 Then LibHasNotes = (Len(gHere(h).Notes) > 0)
End Function

Public Function LibNotesUnread(ByVal key As String) As Boolean
    Dim h As Long
    h = HereIndex(PakOf(key))
    If h >= 0 Then LibNotesUnread = (Len(gHere(h).Notes) > 0 And Not gHere(h).NotesSeen)
End Function

' The first game in a pak, for picking it in the list
Public Function LibFirstGame(ByVal pak As String) As String
    Dim i As Long
    For i = 0 To gShowCount - 1
        If PakOf(gShowId(i)) = PakOf(pak) Then
            LibFirstGame = gShowId(i)
            Exit Function
        End If
    Next
End Function

' How much room installing a game needs: the rest of the download, plus however
' much bigger the game is than the copy it replaces, plus a little to spare
Private Function SpaceNeeded(ByVal c As Long, ByVal target As String) As Double
    Dim dl As String, k As Long, got As Double, old As Double, h As Long, grow As Double
    dl = DlFolder(c, target)
    If ReadText(dl & "\pak.txt") = gPak(c).Made & "|" & gPak(c).ArcBytes Then
        For k = 1 To gPak(c).Pieces
            got = got + FileBytes(dl & "\" & PieceName(gPak(c).Archive, k))
        Next
    End If
    h = HereIndex(gPak(c).Id)
    If h >= 0 Then old = gHere(h).Size
    grow = gPak(c).Size - old
    If grow < 0 Then grow = 0
    SpaceNeeded = (gPak(c).ArcBytes - got) + grow + SPARE_BYTES
End Function

' The download folder sits next to the game: C:\LANPartyGames\_downloads\quake-v1.09
Private Function DlFolder(ByVal c As Long, ByVal target As String) As String
    Dim p As String, i As Long, s As String, ch As String
    p = gPak(c).Path
    If Right$(p, 1) = "/" Then p = Left$(p, Len(p) - 1)
    p = Mid$(p, InStrRev(p, "/") + 1)
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then
            s = s & ch
        Else
            s = s & "_"
        End If
    Next
    If Len(s) = 0 Or s = "." Or s = ".." Then s = gPak(c).Id
    DlFolder = ParentOf(target) & "\_downloads\" & s
End Function

Private Function PieceName(ByVal archive As String, ByVal k As Long) As String
    PieceName = archive & "." & Format$(k, "000")
End Function

' ================================================================ asking to install, update, uninstall
' These take a game's key, or a pak's id. Installing a game installs the whole pak it's in.

' Install or update a game. Asks first when there's something to warn about. True if it's on its way.
Public Function LibAskInstall(ByVal key As String) As Boolean
    Dim c As Long, h As Long, id As String, target As String, need As Double, free As Double, missing As String, s As String
    id = PakOf(key)
    c = CatIndex(id)
    If c < 0 Then
        MsgBox "This game is no longer available on the Cartridge Server.", vbExclamation, "Game Library"
        Exit Function
    End If
    If PakBusy(id) Then Exit Function
    missing = ToolsMissing()
    If Len(missing) > 0 Then
        MsgBox "The following files required by the Game Library are missing: " & missing & "." & vbCrLf & vbCrLf & _
               "Reinstall Palms LAN Center to put them back.", vbExclamation, "Game Library"
        Exit Function
    End If
    If Len(gPakBase) = 0 Then
        MsgBox "The Cartridge Server address has not been received from the Palms Server. Unable to download.", vbExclamation, "Game Library"
        Exit Function
    End If
    target = LibPlace(id)
    If Not (target Like "[A-Za-z]:\*") Then
        MsgBox "The games folder (" & GamesRoot() & ") is not a valid folder on this PC. Select another folder in Tools > Options.", vbExclamation, "Game Library"
        Exit Function
    End If
    h = HereIndex(id)
    If h >= 0 Then
        s = "Updating " & gPak(c).Name & " reinstalls the game from the new cartridge." & vbCrLf & vbCrLf & _
            "All files in " & target & " will be replaced, including configuration files and saved games. " & _
            "Back up any files you want to keep before continuing." & vbCrLf & vbCrLf & _
            "The game can still be played while the update downloads. Continue?"
        If MsgBox(s, vbOKCancel + vbExclamation + vbDefaultButton2, "Update " & gPak(c).Name) <> vbOK Then Exit Function
    ElseIf FolderThere(target) Then
        s = "The folder " & target & " already exists. It may be left over from an incomplete installation." & vbCrLf & vbCrLf & _
            "Installing " & gPak(c).Name & " will replace its contents. Continue?"
        If MsgBox(s, vbOKCancel + vbExclamation + vbDefaultButton2, "Install " & gPak(c).Name) <> vbOK Then Exit Function
    End If
    need = SpaceNeeded(c, target)
    free = FreeBytes(target)
    If free >= 0 And free < need Then
        MsgBox gPak(c).Name & " requires " & SizeText(need) & " of free space on drive " & Left$(target, 2) & ". Available: " & SizeText(free) & "." & vbCrLf & vbCrLf & _
               "Free some disk space, or select a games folder on another drive in Tools > Options.", vbExclamation, "Game Library"
        Exit Function
    End If
    If id = gFailId Then gFailId = ""
    QueueAdd "i" & id
    LibAskInstall = True
End Function

Public Function LibAskUninstall(ByVal key As String) As Boolean
    Dim h As Long, id As String, s As String
    id = PakOf(key)
    h = HereIndex(id)
    If h < 0 Then Exit Function
    If PakBusy(id) Then Exit Function
    If LibPackCount(id) > 1 Then
        s = "Uninstall " & gHere(h).Name & "?" & vbCrLf & vbCrLf & _
            "All games on this multicart will be uninstalled: " & PackNames(id) & "." & vbCrLf & vbCrLf & _
            "The folder " & gHere(h).Place & " will be deleted, including any configuration files and saved games."
    Else
        s = "Uninstall " & gHere(h).Name & "?" & vbCrLf & vbCrLf & _
            "The folder " & gHere(h).Place & " will be deleted, including any configuration files and saved games."
    End If
    If MsgBox(s, vbOKCancel + vbQuestion + vbDefaultButton2, "Uninstall " & gHere(h).Name) <> vbOK Then Exit Function
    If id = gFailId Then gFailId = ""
    QueueAdd "u" & id
    LibAskUninstall = True
End Function

' Every game in a pak, like "Half-Life, Counter-Strike 1.6 and Opposing Force"
Private Function PackNames(ByVal pak As String) As String
    Dim i As Long, n As Long, total As Long, s As String
    total = LibPackCount(pak)
    For i = 0 To gShowCount - 1
        If PakOf(gShowId(i)) = pak Then
            n = n + 1
            If n > 1 And n = total Then
                s = s & " and "
            ElseIf n > 1 Then
                s = s & ", "
            End If
            s = s & LibName(gShowId(i))
        End If
    Next
    PackNames = s
End Function

' Stop installing (or take it out of the queue). True if anything changed.
Public Function LibAskCancel(ByVal key As String) As Boolean
    Dim i As Long, s As String, id As String
    id = PakOf(key)
    i = QueuePos(id)
    If i >= 0 Then
        QueueRemove i
        LibAskCancel = True
        Exit Function
    End If
    If id <> gJobId Or Len(id) = 0 Then Exit Function
    If mRemoving Or mStep = ST_CLEAR Or mStep = ST_REG Or mStep = ST_WALL Then
        MsgBox "This step cannot be cancelled. It will finish in a few seconds.", vbInformation, "Game Library"
        Exit Function
    End If
    If mStep = ST_UNPACK Then
        s = "Cancel the installation of " & mPak.Name & "?" & vbCrLf & vbCrLf & "Files extracted so far will be removed. Downloaded files are kept and are not downloaded again next time."
    Else
        s = "Cancel the download of " & mPak.Name & "?" & vbCrLf & vbCrLf & "Files downloaded so far are kept. The download resumes the next time the game is installed."
    End If
    If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2, "Game Library") <> vbYes Then Exit Function
    If id <> gJobId Then Exit Function          ' it finished while we were asking
    StopJob True
    LibAskCancel = True
End Function

' ================================================================ playing

' Starts the game (or its settings program), in its folder like a shortcut would
Public Sub LibPlay(ByVal key As String, ByVal hwnd As Long, ByVal settingsProgram As Boolean)
    Dim h As Long, g As Long, exe As String, args As String, path As String, wd As String, r As Long, nm As String
    g = HereGameIndex(key)
    h = HereIndex(PakOf(key))
    If g < 0 Or h < 0 Then
        If h >= 0 Then MsgBox "The installed version of " & gHere(h).Name & " does not include this game. Update " & gHere(h).Name & " to install it.", vbInformation, "Game Library"
        Exit Sub
    End If
    nm = mHereGame(g).Name
    If settingsProgram Then
        exe = mHereGame(g).CfgExe
        args = mHereGame(g).CfgArgs
    Else
        exe = mHereGame(g).Exe
        args = mHereGame(g).Args
    End If
    If Len(exe) = 0 Then Exit Sub
    path = JoinPath(gHere(h).Place, exe)
    If Not FileThere(path) Then
        MsgBox "Cannot find " & path & "." & vbCrLf & vbCrLf & "The game files may have been moved or deleted. Uninstall the game and install it again.", vbExclamation, nm
        Exit Sub
    End If
    ' the extra steps, the first time
    If Not settingsProgram And LibNotesUnread(key) Then
        gHere(h).NotesSeen = True
        PIniSet gHere(h).Id, "notes_seen", "1"
        If MsgBox(gHere(h).Name & " has setup notes that should be read before playing." & vbCrLf & vbCrLf & "View them now?", vbYesNo + vbQuestion, nm) = vbYes Then
            LibShowNotes key
            Exit Sub
        End If
    End If
    ' where it starts: its own folder, unless the pak says otherwise (a mod's folder, say)
    wd = ParentOf(path)
    If Len(mHereGame(g).StartIn) > 0 And Not settingsProgram Then
        If FolderThere(JoinPath(gHere(h).Place, mHereGame(g).StartIn)) Then wd = JoinPath(gHere(h).Place, mHereGame(g).StartIn)
    End If
    r = ShellExecute(hwnd, "open", path, args, wd, SW_SHOWNORMAL)
    If r <= 32 Then MsgBox "Unable to start " & path & " (error " & r & ").", vbExclamation, nm
End Sub

Public Sub LibShowNotes(ByVal key As String)
    Dim h As Long
    h = HereIndex(PakOf(key))
    If h < 0 Then Exit Sub
    If Len(gHere(h).Notes) = 0 Then Exit Sub
    If Not gHere(h).NotesSeen Then
        gHere(h).NotesSeen = True
        PIniSet gHere(h).Id, "notes_seen", "1"
    End If
    If Not FileThere(gHere(h).Notes) Then
        MsgBox "The setup notes file cannot be found (" & gHere(h).Notes & ").", vbExclamation, gHere(h).Name
        Exit Sub
    End If
    On Error Resume Next
    Shell "notepad.exe " & Quoted(gHere(h).Notes), vbNormalFocus
End Sub

Public Sub LibOpenFolder(ByVal key As String, ByVal hwnd As Long)
    Dim h As Long
    h = HereIndex(PakOf(key))
    If h < 0 Then Exit Sub
    If FolderThere(gHere(h).Place) Then ShellExecute hwnd, "open", gHere(h).Place, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

' ================================================================ Start menu shortcuts
' Each game gets one in Start > Programs > Palms Arcade Games, so it can be started without
' the LAN Center too. Made with Windows Script Host, which XP has; if it's missing, no shortcuts.

Private Function MenuFolder() As String
    Dim sh As Object
    On Error GoTo Nope
    Set sh = CreateObject("WScript.Shell")
    MenuFolder = sh.SpecialFolders("Programs") & "\" & MENU_NAME
Nope:
End Function

' A shortcut for each game in the pak just installed. Returns where they went, separated by ;
Private Function MakeShortcuts(ByVal place As String) As String
    Dim sh As Object, lnk As Object, menu As String, k As Long, path As String, nm As String, ico As String, out As String
    Dim exe As String, wd As String
    On Error GoTo Done
    menu = MenuFolder()
    If Len(menu) = 0 Then Exit Function
    MakeDirs menu
    Set sh = CreateObject("WScript.Shell")
    For k = 0 To mPakGameCount - 1
        With mPakGames(k)
            nm = FileNameOf(.Name)
            path = menu & "\" & nm & ".lnk"
            ' another installed pak's shortcut with the same name: keep both
            If ShortcutTaken(path) Then path = menu & "\" & nm & " (" & FileNameOf(mPak.Name) & ").lnk"
            exe = JoinPath(place, .Exe)
            wd = ParentOf(exe)
            If Len(.StartIn) > 0 Then wd = JoinPath(place, .StartIn)
            ico = JoinPath(place, "_palms\" & .Id & ".ico")
            If Not FileThere(ico) Then ico = exe
            Set lnk = sh.CreateShortcut(path)
            lnk.TargetPath = exe
            lnk.Arguments = .Args
            lnk.WorkingDirectory = wd
            lnk.IconLocation = ico & ",0"
            lnk.Description = .Name & " (Palms Arcade)"
            lnk.Save
            If FileThere(path) Then out = out & ";" & path
        End With
    Next
Done:
    If Len(out) > 0 Then MakeShortcuts = Mid$(out, 2)
End Function

' A pak's shortcuts, gone
Private Sub RemoveShortcuts(ByVal pak As String)
    Dim h As Long, lnks() As String, k As Long
    h = HereIndex(pak)
    If h < 0 Then Exit Sub
    lnks = Split(gHere(h).Shortcuts, ";")
    For k = 0 To UBound(lnks)
        If LCase$(Right$(lnks(k), 4)) = ".lnk" Then KillQuiet lnks(k)
    Next
    gHere(h).Shortcuts = ""
    If Len(MenuFolder()) > 0 Then RmDirQuiet MenuFolder()      ' only goes if it's empty
End Sub

' A shortcut another installed pak made
Private Function ShortcutTaken(ByVal path As String) As Boolean
    Dim i As Long
    For i = 0 To gHereCount - 1
        If gHere(i).Id <> mPak.Id Then
            If InStr(1, ";" & gHere(i).Shortcuts & ";", ";" & path & ";", vbTextCompare) > 0 Then ShortcutTaken = True
        End If
    Next
End Function

' A game's name as a file name: without the characters Windows won't have in one
Private Function FileNameOf(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("\/:*?""<>|", ch) = 0 Then out = out & ch
    Next
    FileNameOf = Trim$(out)
    If Len(FileNameOf) = 0 Then FileNameOf = "Game"
End Function
Public Function PickFolder(ByVal hwnd As Long, ByVal title As String) As String
    Dim bi As BROWSEINFO, pidl As Long, buf As String, n As Long
    bi.hOwner = hwnd
    bi.pszDisplayName = String$(260, vbNullChar)
    bi.lpszTitle = title
    bi.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    pidl = SHBrowseForFolder(bi)
    If pidl = 0 Then Exit Function
    buf = String$(260, vbNullChar)
    If SHGetPathFromIDList(pidl, buf) <> 0 Then
        n = InStr(buf, vbNullChar)
        If n > 0 Then PickFolder = Left$(buf, n - 1)
    End If
    CoTaskMemFree pidl
End Function

' ================================================================ running programs out of sight

' Starts a program with no window, a little below normal priority so games and the
' LAN Center stay smooth. Its output goes to logFile ("" for none). Returns its
' process handle, or 0 if it couldn't start.
Private Function RunHidden(ByVal cmdLine As String, ByVal logFile As String) As Long
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION, sa As SECURITY_ATTRIBUTES, hLog As Long, ok As Long
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    hLog = INVALID_HANDLE
    If Len(logFile) > 0 Then
        sa.nLength = Len(sa)
        sa.bInheritHandle = 1
        hLog = CreateFileForChild(logFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, sa, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        If hLog <> INVALID_HANDLE Then
            si.dwFlags = si.dwFlags Or STARTF_USESTDHANDLES
            si.hStdOutput = hLog
            si.hStdError = hLog
        End If
    End If
    ok = CreateProcess(vbNullString, cmdLine, 0, 0, 1, CREATE_NO_WINDOW Or BELOW_NORMAL_PRIORITY_CLASS, 0, vbNullString, si, pi)
    ' older Windows don't know those two flags
    If ok = 0 Then ok = CreateProcess(vbNullString, cmdLine, 0, 0, 1, 0, 0, vbNullString, si, pi)
    If hLog <> INVALID_HANDLE Then CloseHandle hLog
    If ok = 0 Then Exit Function
    CloseHandle pi.hThread
    RunHidden = pi.hProcess
End Function

' Starts a program and doesn't wait to see how it goes
Private Sub RunAndForget(ByVal cmdLine As String)
    Dim h As Long
    h = RunHidden(cmdLine, "")
    If h <> 0 Then CloseHandle h
End Sub

' True once the program has finished, with its exit code. Closes the handle when it has.
Private Function ProcDone(ByVal h As Long, code As Long) As Boolean
    If h = 0 Then
        code = -1
        ProcDone = True
        Exit Function
    End If
    If WaitForSingleObject(h, 0) = WAIT_TIMEOUT Then Exit Function
    code = -1
    GetExitCodeProcess h, code
    CloseHandle h
    ProcDone = True
End Function

' Deletes a folder and everything in it. Windows 98 has no rd /s: its deltree does the same.
Private Function DeleteCommand(ByVal folder As String) As String
    Dim cs As String
    cs = Environ$("ComSpec")
    If IsWin9x() Then
        If Len(cs) = 0 Then cs = "command.com"
        DeleteCommand = Quoted(cs) & " /c deltree /y " & Quoted(folder)
    Else
        If Len(cs) = 0 Then cs = "cmd.exe"
        DeleteCommand = Quoted(cs) & " /c rd /s /q " & Quoted(folder)
    End If
End Function

' palmsget.exe, with wget's options: carry on with part-downloaded files, 3 tries, 30 seconds' patience.
' urls: one address a line. They go in a list file, as Windows 98 can't take a long command line.
Private Function DlCommand(ByVal folder As String, ByVal urls As String) As String
    WriteText TempDir() & "\download.txt", urls
    DlCommand = Quoted(AppDir() & "palmsget.exe") & " -c -nv -t 3 -T 30 --no-proxy -o " & Quoted(TempDir() & "\palmsget.log") & _
                " -P " & Quoted(folder) & " -i " & Quoted(TempDir() & "\download.txt")
End Function

' ================================================================ the queue

Private Sub QueueAdd(ByVal job As String)
    If mQueueCount = 0 Then ReDim mQueue(0 To 7)
    If mQueueCount > UBound(mQueue) Then ReDim Preserve mQueue(0 To mQueueCount + 8)
    mQueue(mQueueCount) = job
    mQueueCount = mQueueCount + 1
    gLibDirty = True
    If Len(gJobId) = 0 Then StartNext
    frmMain.LibWake
End Sub

Private Function QueuePos(ByVal id As String) As Long
    Dim i As Long
    QueuePos = -1
    For i = 0 To mQueueCount - 1
        If Mid$(mQueue(i), 2) = id Then
            QueuePos = i
            Exit Function
        End If
    Next
End Function

Private Sub QueueRemove(ByVal i As Long)
    Dim k As Long
    If i < 0 Or i >= mQueueCount Then Exit Sub
    For k = i To mQueueCount - 2
        mQueue(k) = mQueue(k + 1)
    Next
    mQueueCount = mQueueCount - 1
    gLibDirty = True
End Sub

Private Sub StartNext()
    Dim job As String
    Do While Len(gJobId) = 0 And mQueueCount > 0
        job = mQueue(0)
        QueueRemove 0
        If Left$(job, 1) = "u" Then
            BeginRemove Mid$(job, 2)
        Else
            BeginInstall Mid$(job, 2)
        End If
    Loop
End Sub

' True while anything is going on, so the timer keeps looking in
Public Function LibWorking() As Boolean
    LibWorking = (Len(gJobId) > 0 Or mIconProc <> 0 Or mQueueCount > 0)
End Function

Private Sub SetJob(ByVal longText As String, ByVal shortText As String)
    gJobText = longText
    gJobShort = shortText
End Sub

' ================================================================ installing

Private Sub BeginInstall(ByVal id As String)
    Dim c As Long, mark As String, urls As String, k As Long
    c = CatIndex(id)
    If c < 0 Then Exit Sub                   ' gone from the pak server while it waited
    mPak = gPak(c)
    ' its games, as they are now
    ReDim mPakGames(0 To gGameCount)
    mPakGameCount = 0
    For k = 0 To gGameCount - 1
        If gGame(k).PakId = id Then
            mPakGames(mPakGameCount) = gGame(k)
            mPakGameCount = mPakGameCount + 1
        End If
    Next
    If mPakGameCount = 0 Then Exit Sub        ' nothing in it to play
    gJobId = id
    mRemoving = False
    mRegKeys = ""
    mProc = 0
    mTarget = LibPlace(id)
    mDl = DlFolder(c, mTarget)
    gJobPct = 0
    SetJob "Preparing...", "Preparing"
    gLibDirty = True
    MakeDirs mDl
    If Not FolderThere(mDl) Then
        JobFailed "Unable to create the folder " & mDl & ". The drive may be full or read-only."
        Exit Sub
    End If
    ' pieces left over from a different version of this pak would make a mess: start afresh
    mark = mPak.Made & "|" & mPak.ArcBytes
    If ReadText(mDl & "\pak.txt") <> mark Then
        KillQuiet mDl & "\" & mPak.Archive & ".*"
        WriteText mDl & "\pak.txt", mark
    End If
    ' setup.reg and notes.txt first: they're small
    If Len(mPak.Reg) > 0 Or Len(mPak.Notes) > 0 Then
        If Len(mPak.Reg) > 0 Then
            KillQuiet mDl & "\" & mPak.Reg
            urls = urls & gPakBase & mPak.Path & mPak.Reg & vbCrLf
        End If
        If Len(mPak.Notes) > 0 Then
            KillQuiet mDl & "\" & mPak.Notes
            urls = urls & gPakBase & mPak.Path & mPak.Notes & vbCrLf
        End If
        mProc = RunHidden(DlCommand(mDl, urls), "")
        If mProc = 0 Then
            JobFailed "Unable to start palmsget.exe."
            Exit Sub
        End If
        mStep = ST_SMALL
    Else
        StartPieces
    End If
End Sub

Private Sub StartPieces()
    Dim k As Long, urls As String
    For k = 1 To mPak.Pieces
        urls = urls & gPakBase & mPak.Path & PieceName(mPak.Archive, k) & vbCrLf
    Next
    mLastBytes = DownloadedBytes()
    mLastTick = GetTickCount()
    mSpeed = 0
    mProc = RunHidden(DlCommand(mDl, urls), "")
    If mProc = 0 Then
        JobFailed "Unable to start palmsget.exe."
        Exit Sub
    End If
    mStep = ST_PIECES
    UpdateDownload
End Sub

Private Function DownloadedBytes() As Double
    Dim k As Long, total As Double
    For k = 1 To mPak.Pieces
        total = total + FileBytes(mDl & "\" & PieceName(mPak.Archive, k))
    Next
    DownloadedBytes = total
End Function

Private Sub UpdateDownload()
    Dim got As Double, now As Long, dt As Long, inst As Double, pct As Long, s As String
    got = DownloadedBytes()
    now = GetTickCount()
    dt = now - mLastTick
    If dt >= 1500 Then
        inst = (got - mLastBytes) * 1000# / dt
        If inst < 0 Then inst = 0
        If mSpeed = 0 Then
            mSpeed = inst
        Else
            mSpeed = mSpeed * 0.6 + inst * 0.4
        End If
        mLastBytes = got
        mLastTick = now
    End If
    If mPak.ArcBytes > 0 Then
        pct = Int(got / mPak.ArcBytes * 100)
    End If
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    gJobPct = pct
    s = "Downloading cartridge: " & SizeText(got) & " of " & SizeText(mPak.ArcBytes)
    If mSpeed > 0 Then s = s & " at " & Format$(mSpeed / 1048576#, "0.0") & " MB/s"
    SetJob s, "Downloading " & pct & "%"
End Sub

' After the download: clear out the old copy of the game, if there is one
Private Sub StartClear()
    If FolderThere(mTarget) Then
        If Not SafeToDelete(mTarget) Then
            JobFailed "Cannot install to " & mTarget & ". The folder is protected."
            Exit Sub
        End If
        mProc = RunHidden(DeleteCommand(mTarget), "")
        If mProc = 0 Then
            JobFailed "Unable to remove the previous version from " & mTarget & "."
            Exit Sub
        End If
        gJobPct = -1
        SetJob "Removing previous version...", "Removing previous version"
        mStep = ST_CLEAR
    Else
        StartUnpack
    End If
End Sub

Private Sub StartUnpack()
    Dim cmd As String
    mFreeAtStart = FreeBytes(mTarget)
    cmd = Quoted(AppDir() & "7za.exe") & " x -y -o" & Quoted(mTarget) & " " & Quoted(mDl & "\" & PieceName(mPak.Archive, 1))
    mProc = RunHidden(cmd, TempDir() & "\7za.log")
    If mProc = 0 Then
        JobFailed "Unable to start 7za.exe."
        Exit Sub
    End If
    mStep = ST_UNPACK
    gJobPct = 0
    SetJob "Extracting...", "Extracting 0%"
End Sub

' How far along unpacking is: what's gone from the drive since it started
Private Sub UpdateUnpack()
    Dim free As Double, done As Double, pct As Long
    free = FreeBytes(mTarget)
    If free >= 0 And mFreeAtStart > 0 Then done = mFreeAtStart - free
    If done < 0 Then done = 0
    If mPak.Size > 0 Then pct = Int(done / mPak.Size * 100)
    If pct > 99 Then pct = 99
    If done > mPak.Size Then done = mPak.Size
    gJobPct = pct
    SetJob "Extracting: " & SizeText(done) & " of " & SizeText(mPak.Size), "Extracting " & pct & "%"
End Sub

Private Sub StartReg()
    Dim text As String, uni As Boolean
    mRegWarn = ""
    mRegCheck = ""
    If Len(mPak.Reg) = 0 Then
        StartWall
        Exit Sub
    End If
    ' whatever Windows the file came from, make it one regedit here will take
    text = PrepareReg(ReadRegFile(mDl & "\" & mPak.Reg), uni)
    ' the Pak Maker marked where the game's folder goes
    text = Replace(text, "%INSTALLDIR_REG%", Replace(mTarget, "\", "\\"), , , vbTextCompare)
    text = Replace(text, "%INSTALLDIR%", mTarget, , , vbTextCompare)
    mRegCheck = KeysWithSettings(text)
    mRegKeys = RegKeysOf(text)
    WriteRegFile TempDir() & "\setup.reg", text, uni
    mProc = RunHidden(Quoted(RegeditPath()) & " /s " & Quoted(TempDir() & "\setup.reg"), "")
    If mProc = 0 Then
        JobFailed "Unable to start Registry Editor to apply the game's registry settings."
        Exit Sub
    End If
    gJobPct = -1
    SetJob "Applying registry settings...", "Configuring"
    mStep = ST_REG
End Sub

' Let the game through the Windows XP firewall: each of its programs, one at a time
' (a pack's games often share one, like hl.exe). Windows 98 has no netsh, and that's fine.
Private Sub StartWall()
    Dim k As Long
    ReDim mWall(0 To mPakGameCount)
    mWallCount = 0
    mWallIdx = 0
    For k = 0 To mPakGameCount - 1
        AddWall mPakGames(k).Exe
    Next
    WallNext
End Sub

Private Sub AddWall(ByVal exe As String)
    Dim i As Long
    If Len(exe) = 0 Then Exit Sub
    For i = 0 To mWallCount - 1
        If LCase$(mWall(i)) = LCase$(exe) Then Exit Sub
    Next
    If mWallCount > UBound(mWall) Then ReDim Preserve mWall(0 To mWallCount + 4)
    mWall(mWallCount) = exe
    mWallCount = mWallCount + 1
End Sub

Private Sub WallNext()
    Dim cmd As String
    If mWallIdx >= mWallCount Then
        FinishInstall
        Exit Sub
    End If
    cmd = "netsh firewall add allowedprogram program=" & Quoted(JoinPath(mTarget, mWall(mWallIdx))) & _
          " name=" & Quoted(Replace(mPak.Name, """", "'")) & " mode=ENABLE"
    mWallIdx = mWallIdx + 1
    mProc = RunHidden(cmd, "")
    If mProc = 0 Then
        FinishInstall
        Exit Sub
    End If
    gJobPct = -1
    SetJob "Adding Windows Firewall exception...", "Configuring"
    mStep = ST_WALL
End Sub

Private Sub FinishInstall()
    Dim h As GameHere, id As String, nm As String, k As Long, ico As String
    h.Id = mPak.Id
    h.Name = mPak.Name
    h.Ver = mPak.Ver
    h.Rev = mPak.Rev
    h.Made = mPak.Made
    h.Place = mTarget
    h.Exe = mPak.Exe
    h.Args = mPak.Args
    h.CfgExe = mPak.CfgExe
    h.CfgArgs = mPak.CfgArgs
    h.Size = mPak.Size
    h.RegKeys = mRegKeys
    h.Icon = IconFileOf(mPak.Icon)
    h.Installed = Format$(Now, "yyyy-mm-dd hh:nn")
    For k = 0 To mPakGameCount - 1
        If k > 0 Then h.Games = h.Games & ","
        h.Games = h.Games & mPakGames(k).Id
    Next
    For k = 0 To mWallCount - 1
        If k > 0 Then h.Programs = h.Programs & ";"
        h.Programs = h.Programs & mWall(k)
    Next
    ' each game's icon goes in the game's folder too, for its shortcut
    For k = 0 To mPakGameCount - 1
        ico = mPakGames(k).Icon
        If LCase$(Right$(ico, 4)) = ".bmp" Then
            ico = IconDir() & "\" & Left$(ico, Len(ico) - 4) & ".ico"
            If FileThere(ico) Then
                MakeDirs JoinPath(mTarget, "_palms")
                On Error Resume Next
                FileCopy ico, JoinPath(mTarget, "_palms\" & mPakGames(k).Id & ".ico")
                On Error GoTo 0
            End If
        End If
    Next
    ' Start menu shortcuts
    h.Shortcuts = MakeShortcuts(mTarget)
    ' the extra steps live in the game's folder from now on
    If Len(mPak.Notes) > 0 Then
        On Error Resume Next
        FileCopy mDl & "\" & mPak.Notes, JoinPath(mTarget, NOTES_NAME)
        On Error GoTo 0
        If FileThere(JoinPath(mTarget, NOTES_NAME)) Then h.Notes = JoinPath(mTarget, NOTES_NAME)
    End If
    PutHere h
    ' tidy up the download
    For k = 1 To mPak.Pieces
        KillQuiet mDl & "\" & PieceName(mPak.Archive, k)
    Next
    If Len(mPak.Reg) > 0 Then KillQuiet mDl & "\" & mPak.Reg
    If Len(mPak.Notes) > 0 Then KillQuiet mDl & "\" & mPak.Notes
    KillQuiet mDl & "\pak.txt"
    RmDirQuiet mDl
    RmDirQuiet ParentOf(mDl)               ' _downloads, if nothing else is downloading
    id = gJobId
    nm = mPak.Name
    JobDone
    If Len(mRegWarn) > 0 Then
        ' installed, but it may not run: the card and a pop-up say why
        gFailId = id
        gFailText = mRegWarn
        frmMain.LibFinished id, True, nm & " installed, but " & LCase$(Left$(mRegWarn, 1)) & Mid$(mRegWarn, 2), (Len(h.Notes) > 0)
    Else
        frmMain.LibFinished id, True, nm & " installed.", (Len(h.Notes) > 0)
    End If
    StartNext
End Sub

' ================================================================ uninstalling

Private Sub BeginRemove(ByVal id As String)
    Dim h As Long, progs() As String, k As Long
    h = HereIndex(id)
    If h < 0 Then Exit Sub
    gJobId = id
    mRemoving = True
    mProc = 0
    mTarget = gHere(h).Place
    mRegKeys = gHere(h).RegKeys
    mPak.Id = id
    mPak.Name = gHere(h).Name
    mPak.Exe = gHere(h).Exe
    ' its programs, to take off the firewall
    progs = Split(gHere(h).Programs, ";")
    ReDim mWall(0 To UBound(progs) + 1)
    mWallCount = 0
    mWallIdx = 0
    For k = 0 To UBound(progs)
        AddWall Trim$(progs(k))
    Next
    gJobPct = -1
    SetJob "Uninstalling...", "Uninstalling"
    gLibDirty = True
    If FolderThere(mTarget) Then
        If Not SafeToDelete(mTarget) Then
            JobFailed "Cannot delete " & mTarget & ". The folder is protected."
            Exit Sub
        End If
        mProc = RunHidden(DeleteCommand(mTarget), "")
        If mProc = 0 Then
            JobFailed "Unable to delete " & mTarget & "."
            Exit Sub
        End If
        mStep = ST_REMOVE
    Else
        UnwallNext
    End If
End Sub

Private Sub UnwallNext()
    If mWallIdx >= mWallCount Then
        StartUnreg
        Exit Sub
    End If
    mProc = RunHidden("netsh firewall delete allowedprogram program=" & Quoted(JoinPath(mTarget, mWall(mWallIdx))), "")
    mWallIdx = mWallIdx + 1
    If mProc = 0 Then
        StartUnreg
        Exit Sub
    End If
    mStep = ST_UNWALL
End Sub

Private Sub StartUnreg()
    Dim keys() As String, i As Long, text As String
    If Len(mRegKeys) = 0 Then
        FinishRemove
        Exit Sub
    End If
    keys = Split(mRegKeys, ";")
    text = "REGEDIT4" & vbCrLf & vbCrLf
    For i = 0 To UBound(keys)
        If Len(Trim$(keys(i))) > 0 Then text = text & "[-" & Trim$(keys(i)) & "]" & vbCrLf & vbCrLf
    Next
    WriteText TempDir() & "\remove.reg", text
    mProc = RunHidden(Quoted(RegeditPath()) & " /s " & Quoted(TempDir() & "\remove.reg"), "")
    If mProc = 0 Then
        FinishRemove
        Exit Sub
    End If
    mStep = ST_UNREG
End Sub

Private Sub FinishRemove()
    Dim id As String, nm As String
    id = gJobId
    nm = mPak.Name
    RemoveShortcuts id
    DropHere id
    JobDone
    frmMain.LibFinished id, True, nm & " uninstalled.", False
    StartNext
End Sub

' ================================================================ registry files
' Windows 98 writes registry files as plain text starting REGEDIT4. XP and Vista write
' "Windows Registry Editor Version 5.00" files in Unicode. A file without its first line
' is ignored by regedit altogether. The Pak Maker cleans every setup.reg up into
' REGEDIT4, but a pak made by hand might not be, so it's checked again here.

' A registry file as text, whichever way it was saved: Unicode, UTF-8 or plain
Private Function ReadRegFile(ByVal path As String) As String
    Dim fnum As Integer, n As Long, b() As Byte, s As String, w As Long
    On Error GoTo Nope
    fnum = FreeFile
    Open path For Binary Access Read Shared As #fnum
    n = LOF(fnum)
    If n = 0 Then
        Close #fnum
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #fnum, , b
    Close #fnum
    If n >= 2 And b(0) = &HFF And b(1) = &HFE Then
        ' Unicode, from XP or Vista's regedit: the bytes are the string already
        s = b
        If Len(s) > 0 Then s = Mid$(s, 2)
    ElseIf n >= 3 And b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
        ' UTF-8, from a newer Notepad
        w = MultiByteToWideChar(CP_UTF8, 0, b(3), n - 3, 0, 0)
        If w > 0 Then
            s = String$(w, vbNullChar)
            MultiByteToWideChar CP_UTF8, 0, b(3), n - 3, StrPtr(s), w
        Else
            s = Mid$(StrConv(b, vbUnicode), 4)
        End If
    Else
        s = StrConv(b, vbUnicode)
    End If
    ReadRegFile = s
    Exit Function
Nope:
    On Error Resume Next
    Close #fnum
End Function

' Puts right what would stop regedit taking a file: no first line, short names like HKLM,
' keys under Wow6432Node from a 64-bit PC, and no blank line at the end.
' uni comes back True for a Version 5.00 file, which has to be written as Unicode.
Private Function PrepareReg(ByVal text As String, uni As Boolean) As String
    Dim lines() As String, i As Long, t As String, head As String, found As Boolean
    uni = False
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    Do While Len(text) > 0
        If AscW(Left$(text, 1)) <> &HFEFF Then Exit Do
        text = Mid$(text, 2)
    Loop
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Left$(t, 1) = "[" Then
            ' short names regedit doesn't know
            If UCase$(Left$(t, 6)) = "[HKLM\" Then t = "[HKEY_LOCAL_MACHINE\" & Mid$(t, 7)
            If UCase$(Left$(t, 6)) = "[HKCU\" Then t = "[HKEY_CURRENT_USER\" & Mid$(t, 7)
            If UCase$(Left$(t, 7)) = "[-HKLM\" Then t = "[-HKEY_LOCAL_MACHINE\" & Mid$(t, 8)
            If UCase$(Left$(t, 7)) = "[-HKCU\" Then t = "[-HKEY_CURRENT_USER\" & Mid$(t, 8)
            ' from a 64-bit PC: 32-bit games look without the Wow6432Node part
            t = Replace(t, "\SOFTWARE\Wow6432Node\", "\SOFTWARE\", , , vbTextCompare)
            lines(i) = t
        End If
        If Not found And Len(t) > 0 Then
            found = True
            head = UCase$(t)
        End If
    Next
    text = Join(lines, vbCrLf)
    If head = "REGEDIT4" Then
        ' just right
    ElseIf Left$(head, 33) = "WINDOWS REGISTRY EDITOR VERSION 5" Then
        uni = True
    Else
        ' no first line: regedit would ignore the lot
        text = "REGEDIT4" & vbCrLf & vbCrLf & text
    End If
    ' old regedit needs a line break after the last setting
    Do While Right$(text, 2) = vbCrLf
        text = Left$(text, Len(text) - 2)
    Loop
    PrepareReg = text & vbCrLf & vbCrLf
End Function

Private Sub WriteRegFile(ByVal path As String, ByVal text As String, ByVal uni As Boolean)
    Dim fnum As Integer, b() As Byte
    On Error Resume Next
    Kill path
    On Error GoTo Nope
    If uni Then
        b = ChrW$(&HFEFF) & text
        fnum = FreeFile
        Open path For Binary Access Write As #fnum
        Put #fnum, , b
        Close #fnum
    Else
        WriteText path, text
    End If
    Exit Sub
Nope:
    On Error Resume Next
    Close #fnum
End Sub

' 32-bit games read the 32-bit side of the registry. On 64-bit Windows only the regedit
' in SysWOW64 writes there.
Private Function RegeditPath() As String
    If FileThere(WinDir() & "\SysWOW64\regedit.exe") Then
        RegeditPath = WinDir() & "\SysWOW64\regedit.exe"
    Else
        RegeditPath = WinDir() & "\regedit.exe"
    End If
End Function

' Every key the file puts at least one setting in, separated by ;
' (keys with nothing under them are just the folders above, shared with other games)
Private Function KeysWithSettings(ByVal text As String) As String
    Dim lines() As String, i As Long, t As String, cur As String, s As String
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Left$(t, 1) = "[" Then
            cur = ""
            If Left$(t, 6) = "[HKEY_" And Right$(t, 1) = "]" Then cur = Mid$(t, 2, Len(t) - 2)
        ElseIf Len(cur) > 0 And (Left$(t, 1) = """" Or Left$(t, 1) = "@") Then
            If InStr(1, ";" & s & ";", ";" & cur & ";", vbTextCompare) = 0 Then
                If Len(s) > 0 Then s = s & ";"
                s = s & cur
            End If
            cur = ""
        End If
    Next
    KeysWithSettings = s
End Function

' The keys setup.reg adds that belong to the game alone, so uninstalling can take them
' away again: ones it puts settings in, at least three deep, and never shared places
' like anything under SOFTWARE\Microsoft.
Private Function RegKeysOf(ByVal text As String) As String
    Dim keys() As String, i As Long, k As String, parts() As String, s As String
    keys = Split(KeysWithSettings(text), ";")
    For i = 0 To UBound(keys)
        k = keys(i)
        parts = Split(LCase$(k), "\")
        If UBound(parts) >= 3 Then
            If (parts(0) = "hkey_local_machine" Or parts(0) = "hkey_current_user") And parts(1) = "software" Then
                Select Case parts(2)
                Case "microsoft", "classes", "policies", "clients", "odbc", "wow6432node", "description", "intel", "nvidia corporation", "ati technologies"
                Case Else
                    If Len(s) > 0 Then s = s & ";"
                    s = s & k
                End Select
            End If
        End If
    Next
    RegKeysOf = s
End Function

Private Function RootKey(ByVal name As String) As Long
    Select Case UCase$(name)
    Case "HKEY_LOCAL_MACHINE": RootKey = &H80000002
    Case "HKEY_CURRENT_USER": RootKey = &H80000001
    Case "HKEY_CLASSES_ROOT": RootKey = &H80000000
    Case "HKEY_USERS": RootKey = &H80000003
    Case "HKEY_CURRENT_CONFIG": RootKey = &H80000005
    End Select
End Function

' After regedit: are the game's settings really there? Usually they aren't because
' Windows isn't signed in as an administrator.
Private Sub CheckRegWentIn()
    Dim keys() As String, i As Long, p As Long, root As Long, hk As Long, total As Long, missing As Long
    mRegWarn = ""
    If Len(mRegCheck) = 0 Then
        If Len(mPak.Reg) > 0 Then mRegWarn = "The setup.reg file contains no settings. The cartridge may need to be rebuilt."
        Exit Sub
    End If
    keys = Split(mRegCheck, ";")
    For i = 0 To UBound(keys)
        p = InStr(keys(i), "\")
        If p > 0 Then
            root = RootKey(Left$(keys(i), p - 1))
            If root <> 0 Then
                total = total + 1
                If RegOpenKeyEx(root, Mid$(keys(i), p + 1), 0, KEY_READ, hk) = 0 Then
                    RegCloseKey hk
                Else
                    missing = missing + 1
                End If
            End If
        End If
    Next
    If missing > 0 Then
        mRegWarn = "Registry settings were not applied (" & missing & " of " & total & " keys missing). " & _
                   "Log on to Windows as an administrator, then update or reinstall the game."
    End If
End Sub

' ================================================================ the timer

' Called twice a second by the main window while anything's going on.
' True when the Game Library page needs redrawing.
Public Function LibTick() As Boolean
    Dim code As Long
    If mIconProc <> 0 Then
        If ProcDone(mIconProc, code) Then
            mIconProc = 0
            gIconsNew = True
            LibTick = True
        End If
    End If
    If Len(gJobId) = 0 Then
        If mQueueCount > 0 Then StartNext
        Exit Function
    End If
    LibTick = True
    If mProc <> 0 Then
        If Not ProcDone(mProc, code) Then
            Select Case mStep
            Case ST_PIECES: UpdateDownload
            Case ST_UNPACK: UpdateUnpack
            End Select
            Exit Function
        End If
        mProc = 0
    End If
    StepDone code
End Function

' The program for one step has finished: check how it went, and start the next step
Private Sub StepDone(ByVal code As Long)
    Dim got As Double, k As Long
    Select Case mStep
    Case ST_SMALL
        If code <> 0 Then
            JobFailed DlTrouble(code)
        ElseIf Len(mPak.Reg) > 0 And Not FileThere(mDl & "\" & mPak.Reg) Then
            JobFailed "Unable to download " & mPak.Reg & "."
        Else
            StartPieces
        End If
    Case ST_PIECES
        If code <> 0 Then
            JobFailed DlTrouble(code)
            Exit Sub
        End If
        got = DownloadedBytes()
        If mPak.ArcBytes > 0 And Abs(got - mPak.ArcBytes) > 0.5 Then
            For k = 1 To mPak.Pieces
                KillQuiet mDl & "\" & PieceName(mPak.Archive, k)
            Next
            JobFailed "The downloaded files are the wrong size (" & SizeText(got) & ", expected " & SizeText(mPak.ArcBytes) & "). Install the game again to restart the download."
            Exit Sub
        End If
        StartClear
    Case ST_CLEAR
        If FolderThere(mTarget) Then
            JobFailed "Unable to remove all files in " & mTarget & ". The game may still be running."
            Exit Sub
        End If
        ' the old copy is gone now, and so are its shortcuts
        RemoveShortcuts mPak.Id
        DropHere mPak.Id
        StartUnpack
    Case ST_UNPACK
        If code < 0 Or code > 1 Then
            ' remove the half-unpacked game before anything else starts
            If SafeToDelete(mTarget) Then RunAndForget DeleteCommand(mTarget)
            JobFailed UnpackTrouble(code)
            Exit Sub
        End If
        For k = 0 To mPakGameCount - 1
            If Not FileThere(JoinPath(mTarget, mPakGames(k).Exe)) Then
                JobFailed "The files were extracted, but " & mPakGames(k).Exe & " was not found. The cartridge may have been built incorrectly."
                Exit Sub
            End If
        Next
        StartReg
    Case ST_REG
        CheckRegWentIn
        StartWall
    Case ST_WALL
        WallNext
    Case ST_REMOVE
        If FolderThere(mTarget) Then
            JobFailed "Unable to delete all files in " & mTarget & ". The game may still be running."
            Exit Sub
        End If
        UnwallNext
    Case ST_UNWALL
        UnwallNext
    Case ST_UNREG
        FinishRemove
    Case Else
        JobDone
    End Select
End Sub

' palmsget's exit codes are wget's
Private Function DlTrouble(ByVal code As Long) As String
    Dim last As String
    last = LastLine(ReadTail(TempDir() & "\palmsget.log"))
    Select Case code
    Case 3
        DlTrouble = "Unable to save the download. The drive may be full."
    Case 4
        DlTrouble = "Unable to connect to the Cartridge Server (" & gPakBase & ")."
    Case 8
        DlTrouble = "File not found on the Cartridge Server. The cartridge may have been removed or rebuilt. (" & last & ")"
    Case Else
        DlTrouble = "The download failed (palmsget error " & code & ")."
        If Len(last) > 0 Then DlTrouble = DlTrouble & " " & last
    End Select
End Function

Private Function UnpackTrouble(ByVal code As Long) As String
    Dim lines() As String, i As Long, s As String
    lines = Split(Replace(ReadTail(TempDir() & "\7za.log"), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        If InStr(1, lines(i), "error", vbTextCompare) > 0 Or InStr(1, lines(i), "space", vbTextCompare) > 0 Then
            If Len(Trim$(lines(i))) > 0 And Len(s) = 0 Then s = Trim$(lines(i))
        End If
    Next
    UnpackTrouble = "7-Zip could not extract the archive (error " & code & ")."
    If Len(s) > 0 Then UnpackTrouble = UnpackTrouble & " " & s
End Function

Private Sub JobDone()
    gJobId = ""
    mStep = ST_IDLE
    mProc = 0
    gJobPct = -1
    SetJob "", ""
    gLibDirty = True
End Sub

Private Sub JobFailed(ByVal why As String)
    Dim id As String
    id = gJobId
    If mProc <> 0 Then
        TerminateProcess mProc, 1
        CloseHandle mProc
        mProc = 0
    End If
    gFailId = id
    gFailText = why
    JobDone
    frmMain.LibFinished id, False, why, False
    StartNext
End Sub

' Stops the job straight away. The download is kept for next time; a half-unpacked game is removed.
Private Sub StopJob(ByVal goOn As Boolean)
    If mProc <> 0 Then
        TerminateProcess mProc, 1
        WaitForSingleObject mProc, 1000
        CloseHandle mProc
        mProc = 0
    End If
    If mStep = ST_UNPACK And SafeToDelete(mTarget) Then RunAndForget DeleteCommand(mTarget)
    JobDone
    If goOn Then StartNext
End Sub

' The LAN Center is closing: stop whatever's running
Public Sub LibShutdown()
    mQueueCount = 0
    If Len(gJobId) > 0 Then StopJob False
    If mIconProc <> 0 Then
        TerminateProcess mIconProc, 1
        CloseHandle mIconProc
        mIconProc = 0
    End If
End Sub

' ================================================================ icons

' Fetches the icons we haven't got yet, all in one go, in the background
' (the .bmp for the list, and the .ico for the game's Start menu shortcut)
Public Sub LibFetchIcons()
    Dim i As Long, f As String, list As String, n As Long, ico As String
    If mIconProc <> 0 Then Exit Sub
    If Len(gPakBase) = 0 Then Exit Sub
    If Not FileThere(AppDir() & "palmsget.exe") Then Exit Sub
    For i = 0 To gGameCount - 1
        f = gGame(i).Icon
        If Len(f) > 0 Then
            If Not FileThere(IconDir() & "\" & f) Then
                list = list & gPakBase & "icons/" & f & vbCrLf
                n = n + 1
            End If
            If LCase$(Right$(f, 4)) = ".bmp" Then
                ico = Left$(f, Len(f) - 4) & ".ico"
                ' an older Cartridge Server hasn't any: ask once, not every time
                If Not FileThere(IconDir() & "\" & ico) And InStr(mIcoTried, "|" & ico & "|") = 0 Then
                    mIcoTried = mIcoTried & "|" & ico & "|"
                    list = list & gPakBase & "icons/" & ico & vbCrLf
                    n = n + 1
                End If
            End If
        End If
    Next
    ' the screenshots, for the projector
    For i = 0 To gPakCount - 1
        f = gPak(i).Shot
        If Len(f) > 0 Then
            If Not FileThere(IconDir() & "\" & f) Then
                list = list & gPakBase & "icons/" & f & vbCrLf
                n = n + 1
            End If
        End If
    Next
    If n = 0 Then Exit Sub
    MakeDirs IconDir()
    WriteText TempDir() & "\icons.txt", list
    mIconProc = RunHidden(Quoted(AppDir() & "palmsget.exe") & " -nv -nc -t 2 -T 15 --no-proxy -o " & Quoted(TempDir() & "\icons.log") & _
                          " -P " & Quoted(IconDir()) & " -i " & Quoted(TempDir() & "\icons.txt"), "")
    If mIconProc <> 0 Then frmMain.LibWake
End Sub
