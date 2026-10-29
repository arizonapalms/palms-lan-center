Attribute VB_Name = "modCartridge"
' Palms LAN Center - Makes cartridges: scans a game folder and packs it with 7-Zip for the Game Library.
' Also runs as VBScript, so no $ functions, Format, IIf or Like.
Option Explicit

Private Const GAMES_ROOT As String = "C:\LANPartyGames\"     ' where players' PCs install games, for the long path check

Private fso As Object
Private sh As Object
Private here As String                ' the LAN Center's folder, with 7za.exe in it
Private makerVer As String            ' the LAN Center's version, written into pak.ini
Private serverShare As String         ' the Cartridge Server's upload share: the host approves what arrives
Private oldShare As String            ' the old name of the upload share, still in some pak.ini files
Private libraryShare As String        ' the approved cartridges, read only
Private adminUrl As String            ' where the host approves them
Private pieceMB As Long
Private archiveBytes As Double

Private prev As Object
Private prevPath As String
Private prevGames As Object
Private prevGameCount As Long

Private srcDir As String
Private fileRel() As String
Private fileSize() As Double
Private fileCount As Long
Private totalBytes As Double
Private exeList() As String
Private exeCount As Long
Private regList() As String
Private regCount As Long
Private lnkList() As String
Private lnkCount As Long

Private gName As String
Private gId As String
Private gVer As String
Private gRev As String          ' the cartridge's revision, 1 for the first
Private gExe As String
Private gArgs As String
Private gCfg As String
Private gCfgArgs As String
Private gIcon As String
Private gPlayers As String
Private gFolder As String
Private gReg As String
Private gRegPath As String
Private gNotes As String
Private outRoot As String

Private regKeys As Long
Private regValues As Long
Private regNotes As String
Private regTrouble As String
Private lastRegText As String

Private sharedArgs As String
Private gmIconFile() As String
Private gmIconFrom() As String
Private isPack As Boolean
Private packArgs As String

' games found in the folder (from shortcuts and mod folders)
Private cName() As String
Private cExe() As String
Private cArgs() As String
Private cStart() As String
Private cIcon() As String
Private cPlayers() As String
Private cCfg() As String
Private cCfgArgs() As String
Private cCount As Long

' the games on a multicart, as the wizard sends them
Private gmId() As String
Private gmName() As String
Private gmExe() As String
Private gmArgs() As String
Private gmStart() As String
Private gmIcon() As String
Private gmPlayers() As String
Private gmCfg() As String
Private gmCfgArgs() As String
Private gmCount As Long

' looking through the folder a bit at a time
Private scanPath() As String
Private scanRel() As String
Private scanTop As Long

' making the cartridge
Private buildDir As String
Private buildTmp As String
Private buildStopped As Boolean
Private buildSizes As Object
Private buildIconName As String
Private buildIconFrom As String
Private buildWarn As String
Private logPos As Long
Private logRest As String
Private doneFiles As Long
Private doneBytes As Double
Private newer7z As Boolean

' ---------------------------------------------------------------- starting up

' folder: where 7za.exe is (the LAN Center's folder). ver: the LAN Center's version, for pak.ini.
Public Sub EngInit(ByVal folder As String, ByVal ver As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    here = folder
    If Right(here, 1) = "\" And Len(here) > 3 Then here = Left(here, Len(here) - 1)
    makerVer = ver
    pieceMB = 2000
    EngSetServer "10.0.200.73"
    Set prev = CreateObject("Scripting.Dictionary")
    prev.CompareMode = 1
    Set prevGames = CreateObject("Scripting.Dictionary")
    prevGameCount = 0
    prevPath = ""
    srcDir = ""
    fileCount = 0
    totalBytes = 0
    exeCount = 0
    regCount = 0
    lnkCount = 0
    ReDim fileRel(0)
    ReDim fileSize(0)
    ReDim exeList(0)
    ReDim regList(0)
    ReDim lnkList(0)
    isPack = False
    packArgs = ""
    sharedArgs = ""
    gName = ""
    gId = ""
    gVer = ""
    gRev = "1"
    gPlayers = ""
    StartCandidates
    EngClearGames
End Sub

' The Cartridge Server, from the address the Palms Server gives out (like 10.0.200.73)
Public Sub EngSetServer(ByVal host As String)
    If Len(host) = 0 Then host = "10.0.200.73"
    serverShare = "\\" & host & "\incoming"
    oldShare = "\\" & host & "\paks"
    libraryShare = "\\" & host & "\library"
    adminUrl = "http://" & host & "/admin/"
End Sub

Public Function EngVersion() As String
    EngVersion = makerVer
End Function

Public Function EngHere() As String
    EngHere = here
End Function

Public Function Eng7zaThere() As Boolean
    Eng7zaThere = fso.FileExists(here & "\7za.exe")
End Function

' For testing: smaller pieces
Public Sub EngSetPieceMB(ByVal n As Long)
    pieceMB = n
End Sub

Public Function EngGamesRoot() As String
    EngGamesRoot = GAMES_ROOT
End Function

Public Function EngServerShare() As String
    EngServerShare = serverShare
End Function

Public Function EngOldShare() As String
    EngOldShare = oldShare
End Function

Public Function EngLibraryShare() As String
    EngLibraryShare = libraryShare
End Function

Public Function EngAdminUrl() As String
    EngAdminUrl = adminUrl
End Function

' Opens a web page in the PC's browser
Public Sub EngOpenUrl(ByVal u As String)
    On Error Resume Next
    sh.Run """" & u & """", 1, False
End Sub

' An approved cartridge with this folder name is already in the Game Library
Public Function EngLibraryHas(ByVal folderName As String) As Boolean
    EngLibraryHas = False
    If Len(folderName) = 0 Then Exit Function
    On Error Resume Next
    EngLibraryHas = fso.FolderExists(libraryShare & "\" & folderName)
End Function

' The Cartridge Server's upload share when it can be reached, otherwise a Cartridges folder here
Public Function EngDefaultOutput() As String
    On Error Resume Next
    If fso.FolderExists(serverShare) Then
        EngDefaultOutput = serverShare
    Else
        EngDefaultOutput = here & "\Cartridges"
    End If
End Function

' Saving straight into the Cartridge Server's upload share
Public Function EngOnServer(ByVal outPath As String) As Boolean
    EngOnServer = False
    If Len(outPath) = 0 Then Exit Function
    EngOnServer = (InStr(1, outPath, serverShare, vbTextCompare) = 1)
End Function

' ---------------------------------------------------------------- files and folders

' A folder picker. Returns the folder, or "" if they cancelled. With a root folder, it
' shows just that folder and the ones inside it (for picking a folder inside the game).
Public Function EngPickFolder(ByVal title As String, ByVal root As String) As String
    Dim shellApp As Object, f As Object, p As String
    EngPickFolder = ""
    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    ' only real folders, with a box to type one in, in the newer style
    If Len(root) > 0 And fso.FolderExists(root) Then
        Set f = shellApp.BrowseForFolder(0, title, &H51, CStr(root))
    Else
        Set f = shellApp.BrowseForFolder(0, title, &H51)
    End If
    If Err.Number <> 0 Then Exit Function
    If f Is Nothing Then Exit Function
    p = f.Self.Path
    If Err.Number <> 0 Then Exit Function
    If fso.FolderExists(p) Then EngPickFolder = p
End Function

Public Function EngFileExists(ByVal p As String) As Boolean
    EngFileExists = False
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    EngFileExists = fso.FileExists(p)
End Function

Public Function EngFolderExists(ByVal p As String) As Boolean
    EngFolderExists = False
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    EngFolderExists = fso.FolderExists(p)
End Function

' A file given as a full path, or as a path inside the game folder. Returns its full path, or "".
Public Function EngResolve(ByVal p As String) As String
    EngResolve = ""
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    If fso.FileExists(p) Then
        EngResolve = fso.GetFile(p).Path
    ElseIf Len(srcDir) > 0 Then
        If fso.FileExists(srcDir & "\" & RelPath(p)) Then EngResolve = srcDir & "\" & RelPath(p)
    End If
End Function

' A file inside the game folder (a path like System\UnrealTournament.exe, or a full path into it)
Public Function EngInGameFile(ByVal p As String) As Boolean
    EngInGameFile = False
    If Len(p) = 0 Or Len(srcDir) = 0 Then Exit Function
    On Error Resume Next
    EngInGameFile = fso.FileExists(srcDir & "\" & RelPath(p))
End Function

Public Function EngInGameFolder(ByVal p As String) As Boolean
    EngInGameFolder = False
    If Len(p) = 0 Or Len(srcDir) = 0 Then Exit Function
    On Error Resume Next
    EngInGameFolder = fso.FolderExists(srcDir & "\" & RelPath(p))
End Function

' A path inside the game folder, without the game folder on the front (or as it was, if it's elsewhere)
Public Function EngRelPath(ByVal p As String) As String
    EngRelPath = RelPath(p)
End Function

' Something a picture can come from: a picture, or a program or DLL to take one out of
Public Function EngIsIconFile(ByVal p As String) As Boolean
    Dim ext As String
    ext = LCase(fso.GetExtensionName(p))
    EngIsIconFile = (InStr(" ico bmp gif jpg jpeg png exe dll ", " " & ext & " ") > 0)
End Function

Public Function EngMakeFolder(ByVal p As String) As Boolean
    EngMakeFolder = MakeFolder(p)
End Function

Public Sub EngOpenFolder(ByVal p As String)
    On Error Resume Next
    sh.Run "explorer.exe """ & p & """", 1, False
End Sub

Public Sub EngOpenFile(ByVal p As String)
    On Error Resume Next
    sh.Run "notepad.exe """ & p & """", 1, False
End Sub

' Free space where the cartridge's going, in bytes, or -1 if Windows can't say
Public Function EngFreeSpace(ByVal p As String) As Double
    EngFreeSpace = -1
    On Error Resume Next
    EngFreeSpace = CDbl(fso.GetDrive(fso.GetDriveName(p)).FreeSpace)
    If Err.Number <> 0 Then EngFreeSpace = -1
End Function

' The folder a file is in
Public Function EngParentOf(ByVal p As String) As String
    On Error Resume Next
    EngParentOf = fso.GetParentFolderName(p)
End Function

' ---------------------------------------------------------------- small helpers for the window

Public Function EngMakeId(ByVal s As String) As String
    EngMakeId = MakeId(s)
End Function

Public Function EngIsId(ByVal s As String) As Boolean
    EngIsId = IsId(s)
End Function

Public Function EngIsPlayers(ByVal s As String) As Boolean
    EngIsPlayers = IsPlayers(s)
End Function

Public Function EngMakeFolderName(ByVal s As String) As String
    EngMakeFolderName = MakeFolderName(s)
End Function

Public Function EngTidy(ByVal s As String) As String
    EngTidy = Tidy(s)
End Function

Public Function EngSizeText(ByVal n As Double) As String
    EngSizeText = SizeText2(n)
End Function

' ---------------------------------------------------------------- looking through the game folder
' A bit at a time, so the window stays awake on a big game over the network.

Public Function EngScanStart(ByVal folder As String) As Boolean
    EngScanStart = False
    If Right(folder, 1) = "\" And Len(folder) > 3 Then folder = Left(folder, Len(folder) - 1)
    If Len(folder) = 0 Then Exit Function
    On Error Resume Next
    If Not fso.FolderExists(folder) Then Exit Function
    srcDir = fso.GetFolder(folder).Path
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    fileCount = 0
    totalBytes = 0
    exeCount = 0
    regCount = 0
    lnkCount = 0
    ReDim fileRel(1000)
    ReDim fileSize(1000)
    ReDim exeList(50)
    ReDim regList(20)
    ReDim lnkList(50)
    ReDim scanPath(50)
    ReDim scanRel(50)
    scanTop = 0
    PushScan srcDir, ""
    EngScanStart = True
End Function

Private Sub PushScan(ByVal path As String, ByVal rel As String)
    If scanTop > UBound(scanPath) Then
        ReDim Preserve scanPath(scanTop * 2)
        ReDim Preserve scanRel(scanTop * 2)
    End If
    scanPath(scanTop) = path
    scanRel(scanTop) = rel
    scanTop = scanTop + 1
End Sub

' Looks through some more of the folder. True once it's all done.
Public Function EngScanStep() As Boolean
    Dim t0 As Double, el As Double, path As String, rel As String
    t0 = Timer
    Do While scanTop > 0
        scanTop = scanTop - 1
        path = scanPath(scanTop)
        rel = scanRel(scanTop)
        ScanOne path, rel
        el = Timer - t0
        If el < 0 Then el = el + 86400
        If el > 0.25 Then Exit Do
    Loop
    EngScanStep = (scanTop = 0)
End Function

' One folder: its files, and its subfolders onto the pile (last first, so they come off in order)
Private Sub ScanOne(ByVal path As String, ByVal rel As String)
    Dim folder As Object, f As Object, r As String, ext As String, subs() As String, n As Long, k As Long
    On Error Resume Next
    Set folder = fso.GetFolder(path)
    If Err.Number <> 0 Then Exit Sub
    For Each f In folder.Files
        r = rel & f.Name
        If LCase(f.Name) <> "thumbs.db" Then
            If fileCount > UBound(fileRel) Then
                ReDim Preserve fileRel(fileCount * 2)
                ReDim Preserve fileSize(fileCount * 2)
            End If
            fileRel(fileCount) = r
            fileSize(fileCount) = CDbl(f.Size)
            fileCount = fileCount + 1
            totalBytes = totalBytes + CDbl(f.Size)
            ext = LCase(fso.GetExtensionName(f.Name))
            ' programs in the top two levels, like Quake\glquake.exe or UT\System\UnrealTournament.exe
            If ext = "exe" And InStr(rel, "\") = InStrRev(rel, "\") Then
                If exeCount > UBound(exeList) Then ReDim Preserve exeList(exeCount * 2)
                exeList(exeCount) = r
                exeCount = exeCount + 1
            End If
            If ext = "reg" And Len(rel) = 0 And regCount <= UBound(regList) Then
                regList(regCount) = r
                regCount = regCount + 1
            End If
            ' shortcuts in the top two levels: how the games are meant to be started
            If ext = "lnk" And InStr(rel, "\") = InStrRev(rel, "\") And lnkCount <= UBound(lnkList) Then
                lnkList(lnkCount) = r
                lnkCount = lnkCount + 1
            End If
        End If
    Next
    n = 0
    ReDim subs(folder.SubFolders.Count)
    For Each f In folder.SubFolders
        subs(n) = f.Name
        n = n + 1
    Next
    For k = n - 1 To 0 Step -1
        PushScan path & "\" & subs(k), rel & subs(k) & "\"
    Next
End Sub

' What it found: files, size (as text), bytes, exes, regs, lnks, src, pieces
Public Function EngScanInfo(ByVal key As String) As Variant
    Dim n As Double
    Select Case LCase(key)
    Case "files"
        EngScanInfo = fileCount
    Case "size"
        EngScanInfo = SizeText2(totalBytes)
    Case "bytes"
        EngScanInfo = totalBytes
    Case "exes"
        EngScanInfo = exeCount
    Case "regs"
        EngScanInfo = regCount
    Case "lnks"
        EngScanInfo = lnkCount
    Case "src"
        EngScanInfo = srcDir
    Case "pieces"
        ' rounded up: a little over 2000 MB is two pieces
        n = -Int(-(totalBytes + 65536) / (CDbl(pieceMB) * 1048576))
        If n < 1 Then n = 1
        EngScanInfo = n
    Case Else
        EngScanInfo = ""
    End Select
End Function

Public Function EngExe(ByVal i As Long) As String
    EngExe = exeList(i)
End Function

Public Function EngRegFile(ByVal i As Long) As String
    EngRegFile = regList(i)
End Function

' The program most likely to start the game
Public Function EngGuessExe(ByVal nm As String, ByVal id As String) As String
    gName = nm
    gId = id
    EngGuessExe = GuessExe()
End Function

' Files whose path will be too long for Windows once installed: how many, then up to 5 of them
Public Function EngLongPaths(ByVal folderName As String) As String
    Dim i As Long, n As Long, s As String
    n = 0
    s = ""
    For i = 0 To fileCount - 1
        If Len(GAMES_ROOT & folderName & "\" & fileRel(i)) > 250 Then
            If n < 5 Then s = s & vbLf & fileRel(i)
            n = n + 1
        End If
    Next
    EngLongPaths = n & s
End Function

' ---------------------------------------------------------------- the games in it

' Looks for games: shortcuts, and Half-Life style mod folders. Returns how many it found.
Public Function EngDetect() As Long
    gPlayers = ""
    sharedArgs = ""
    DetectGames
    EngDetect = cCount
End Function

Public Function EngCand(ByVal i As Long, ByVal key As String) As String
    Select Case LCase(key)
    Case "name": EngCand = cName(i)
    Case "exe": EngCand = cExe(i)
    Case "args": EngCand = cArgs(i)
    Case "start": EngCand = cStart(i)
    Case "icon": EngCand = cIcon(i)
    Case "players": EngCand = cPlayers(i)
    Case "cfg": EngCand = cCfg(i)
    Case "cfgargs": EngCand = cCfgArgs(i)
    Case Else: EngCand = ""
    End Select
End Function

' What every shortcut had in common besides its own -game, like -nomaster -steam
Public Function EngSharedArgs() As String
    EngSharedArgs = sharedArgs
End Function

' ---------------------------------------------------------------- registry files

' Reads and cleans up a registry file. Returns its full path, or "" if it isn't there.
' Then EngRegInfo says what was in it.
Public Function EngCheckReg(ByVal p As String) As String
    Dim full As String
    EngCheckReg = ""
    lastRegText = ""
    regKeys = 0
    regValues = 0
    regNotes = ""
    regTrouble = ""
    full = EngResolve(p)
    If Len(full) = 0 Then Exit Function
    On Error Resume Next
    lastRegText = CleanReg(ReadAnyText(full))
    If Err.Number <> 0 Then
        regTrouble = vbCrLf & "Unable to read the file: " & Err.Description
        Err.Clear
    End If
    EngCheckReg = full
End Function

' keys, values, notes (what was fixed), trouble (what looks wrong), guess (the install folder it names)
Public Function EngRegInfo(ByVal key As String) As Variant
    Select Case LCase(key)
    Case "keys": EngRegInfo = regKeys
    Case "values": EngRegInfo = regValues
    Case "notes": EngRegInfo = TidyLines(regNotes)
    Case "trouble": EngRegInfo = TidyLines(regTrouble)
    Case "guess": EngRegInfo = PickRegPath(lastRegText)
    Case Else: EngRegInfo = ""
    End Select
End Function

' Whether the folder is named in the registry file last checked
Public Function EngRegMentions(ByVal folder As String) As Boolean
    Dim f As String
    f = folder
    If Right(f, 1) = "\" Then f = Left(f, Len(f) - 1)
    EngRegMentions = (InStr(1, lastRegText, Replace(f, "\", "\\"), vbTextCompare) > 0 Or InStr(1, lastRegText, f, vbTextCompare) > 0)
End Function

' Lines without the indenting the command window version used, one per line
Private Function TidyLines(ByVal s As String) As String
    Dim parts() As String, i As Long, out As String
    out = ""
    parts = Split(Replace(s, vbCr, ""), vbLf)
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & Trim(parts(i))
        End If
    Next
    TidyLines = out
End Function

' ---------------------------------------------------------------- rebuilding a cartridge

Public Function EngLoadPrevious(ByVal path As String) As Boolean
    EngLoadPrevious = False
    prev.RemoveAll
    prevGames.RemoveAll
    prevGameCount = 0
    prevPath = ""
    If Not fso.FileExists(path) Then Exit Function
    On Error Resume Next
    LoadPrevious path
    If Err.Number <> 0 Then
        Err.Clear
        prev.RemoveAll
        prevGames.RemoveAll
        prevGameCount = 0
        Exit Function
    End If
    EngLoadPrevious = (prev.Count > 0)
End Function

Public Function EngPrev(ByVal key As String) As String
    EngPrev = ""
    If prev.Exists(key) Then EngPrev = prev.Item(key)
End Function

Public Function EngPrevGameCount() As Long
    EngPrevGameCount = prevGameCount
End Function

' One of the multicart's games from last time (1 to EngPrevGameCount)
Public Function EngPrevGame(ByVal i As Long, ByVal key As String) As String
    Dim k As Variant
    EngPrevGame = ""
    For Each k In prevGames.Keys
        If CLng(k) = i Then EngPrevGame = GetKey(prevGames(k), LCase(key))
    Next
End Function

' ---------------------------------------------------------------- the answers, from the window

Public Sub EngSet(ByVal key As String, ByVal value As String)
    Select Case LCase(key)
    Case "name": gName = Tidy(value)
    Case "id": gId = LCase(value)
    Case "version": gVer = Tidy(value)
    Case "revision": gRev = CStr(Val(value))
    Case "players": gPlayers = Replace(value, " ", "")
    Case "folder": gFolder = value
    Case "exe": gExe = RelPath(value)
    Case "args": gArgs = Trim(value)
    Case "cfg": gCfg = RelPath(value)
    Case "cfgargs": gCfgArgs = Trim(value)
    Case "icon": gIcon = value
    Case "reg": gReg = value
    Case "regpath"
        gRegPath = value
        If Right(gRegPath, 1) = "\" Then gRegPath = Left(gRegPath, Len(gRegPath) - 1)
    Case "notes": gNotes = value
    Case "output"
        outRoot = value
        If Right(outRoot, 1) = "\" And Len(outRoot) > 3 Then outRoot = Left(outRoot, Len(outRoot) - 1)
    Case "pack": isPack = (value = "1" Or LCase(value) = "true")
    Case "packargs": packArgs = Trim(value)
    End Select
End Sub

Public Sub EngClearGames()
    gmCount = 0
    ReDim gmId(40)
    ReDim gmName(40)
    ReDim gmExe(40)
    ReDim gmArgs(40)
    ReDim gmStart(40)
    ReDim gmIcon(40)
    ReDim gmPlayers(40)
    ReDim gmCfg(40)
    ReDim gmCfgArgs(40)
End Sub

' One game on a multicart. Returns its ID within the cartridge.
Public Function EngAddGame(ByVal nm As String, ByVal pl As String, ByVal ex As String, ByVal ar As String, ByVal st As String, ByVal ic As String, ByVal cf As String, ByVal cfa As String) As String
    Dim id As String, k As Long
    If gmCount > UBound(gmId) Then GrowGames
    nm = Tidy(nm)
    id = MakeId(nm)
    k = 2
    Do While GameIdUsed(id)
        id = Left(MakeId(nm), 21) & k
        k = k + 1
    Loop
    gmId(gmCount) = id
    gmName(gmCount) = nm
    pl = Replace(pl, " ", "")
    If Len(pl) = 0 Then pl = gPlayers
    gmPlayers(gmCount) = pl
    gmExe(gmCount) = RelPath(ex)
    gmArgs(gmCount) = Trim(ar)
    gmStart(gmCount) = RelPath(st)
    gmIcon(gmCount) = ic
    gmCfg(gmCount) = RelPath(cf)
    If Len(gmCfg(gmCount)) > 0 Then
        gmCfgArgs(gmCount) = Trim(cfa)
    Else
        gmCfgArgs(gmCount) = ""
    End If
    gmCount = gmCount + 1
    EngAddGame = id
End Function

Private Function GameIdUsed(ByVal id As String) As Boolean
    Dim i As Long
    GameIdUsed = False
    For i = 0 To gmCount - 1
        If gmId(i) = id Then GameIdUsed = True
    Next
End Function

Private Sub GrowGames()
    Dim n As Long
    n = UBound(gmId) + 20
    ReDim Preserve gmId(n)
    ReDim Preserve gmName(n)
    ReDim Preserve gmExe(n)
    ReDim Preserve gmArgs(n)
    ReDim Preserve gmStart(n)
    ReDim Preserve gmIcon(n)
    ReDim Preserve gmPlayers(n)
    ReDim Preserve gmCfg(n)
    ReDim Preserve gmCfgArgs(n)
End Sub

Public Function EngPakDir() As String
    EngPakDir = PakDir()
End Function

Public Function EngPakExists() As Boolean
    EngPakExists = False
    On Error Resume Next
    EngPakExists = fso.FolderExists(PakDir())
End Function

' ---------------------------------------------------------------- making the cartridge
' The small files go in straight away. 7-Zip then runs out of sight from a little batch
' file, which writes down how it went when it's finished, and the window keeps an eye on it.

' Starts making the cartridge. Returns "" when it's under way, or what went wrong.
Public Function EngBuildStart(ByVal replaceOld As Boolean) As String
    Dim bat As String, logFile As String, doneFile As String, t As Object, cmd As String, i As Long, msg As String
    EngBuildStart = ""
    buildDir = PakDir()
    buildStopped = False
    buildWarn = ""
    If Not fso.FileExists(here & "\7za.exe") Then
        EngBuildStart = "7za.exe was not found in " & here & ". Run palms-update.vbs to install it."
        Exit Function
    End If
    If fso.FolderExists(buildDir) Then
        If Not replaceOld Then
            EngBuildStart = "A cartridge already exists in " & buildDir & "."
            Exit Function
        End If
        On Error Resume Next
        fso.DeleteFolder buildDir, True
        On Error GoTo 0
        If fso.FolderExists(buildDir) Then
            EngBuildStart = "Unable to remove the existing cartridge in " & buildDir & ". A file in it may be in use."
            Exit Function
        End If
    End If
    If Not MakeFolder(buildDir) Then
        EngBuildStart = "Unable to create the folder " & buildDir & "."
        Exit Function
    End If

    ' the small files first
    msg = SmallFiles()
    If Len(msg) > 0 Then
        EngBuildStart = msg
        Exit Function
    End If

    ' then 7-Zip, with nothing on screen
    buildTmp = fso.GetSpecialFolder(2).Path & "\palms-cartridge"     ' in the Temp folder
    MakeFolder buildTmp
    bat = buildTmp & "\pack.cmd"
    logFile = buildTmp & "\7za.log"
    doneFile = buildTmp & "\done.txt"
    On Error Resume Next
    If fso.FileExists(doneFile) Then fso.DeleteFile doneFile, True
    If fso.FileExists(logFile) Then fso.DeleteFile logFile, True
    On Error GoTo 0
    If fso.FileExists(doneFile) Then
        EngBuildStart = "Another cartridge is being made on this PC. Wait for it to finish."
        Exit Function
    End If
    newer7z = SevenZipIsNewer()
    cmd = """" & here & "\7za.exe"" a -t7z -mx0 -v" & pieceMB & "m -y -xr!Thumbs.db """ & buildDir & "\game.7z"" """ & srcDir & "\*"""
    ' newer 7-Zips only list each file if asked; 9.20 does it anyway and doesn't know the option
    If newer7z Then cmd = cmd & " -bb1"
    Set t = fso.CreateTextFile(bat, True, False)
    t.WriteLine "@echo off"
    ' a % in a batch file has to be doubled
    t.WriteLine Replace(cmd, "%", "%%") & " > """ & Replace(logFile, "%", "%%") & """ 2>&1"
    t.WriteLine "echo code %errorlevel% > """ & Replace(doneFile, "%", "%%") & """"
    t.Close
    ' each file's size, to tell how far along 7-Zip is from the files it lists
    Set buildSizes = CreateObject("Scripting.Dictionary")
    buildSizes.CompareMode = 1
    For i = 0 To fileCount - 1
        buildSizes(fileRel(i)) = fileSize(i)
    Next
    logPos = 0
    logRest = ""
    doneFiles = 0
    doneBytes = 0
    sh.Run """" & bat & """", 0, False
End Function

' The icon(s), the registry file and the setup notes. Returns what went wrong, or "".
Private Function SmallFiles() As String
    Dim ext As String, f As String, what As String
    SmallFiles = ""
    buildIconName = ""
    buildIconFrom = ""
    On Error Resume Next
    what = "the icon"
    If Len(gIcon) > 0 Then
        ext = LCase(fso.GetExtensionName(gIcon))
        If ext = "exe" Or ext = "dll" Then
            buildIconFrom = RelPath(gIcon)
            If Not fso.FileExists(srcDir & "\" & buildIconFrom) Then
                ' a program outside the game folder: take a copy so the server can find its icon
                fso.CopyFile gIcon, buildDir & "\icon." & ext
                buildIconFrom = ""
                buildIconName = "icon." & ext
            End If
        Else
            f = gIcon
            If Not fso.FileExists(f) Then f = srcDir & "\" & RelPath(f)
            buildIconName = "icon." & ext
            fso.CopyFile f, buildDir & "\" & buildIconName
        End If
    ElseIf Not isPack Then
        buildIconFrom = gExe
    End If
    If Err.Number <> 0 Then
        SmallFiles = "Unable to copy " & what & " to the cartridge (" & Err.Description & ")."
        Exit Function
    End If
    what = "the games' icons"
    If isPack Then PackIcons buildDir
    If Err.Number <> 0 Then
        SmallFiles = "Unable to copy " & what & " to the cartridge (" & Err.Description & ")."
        Exit Function
    End If
    what = "the registry file"
    If Len(gReg) > 0 Then WriteReg gReg, buildDir & "\setup.reg", gRegPath
    If Err.Number <> 0 Then
        SmallFiles = "Unable to copy " & what & " to the cartridge (" & Err.Description & ")."
        Exit Function
    End If
    what = "the setup notes file"
    If Len(gNotes) > 0 Then fso.CopyFile gNotes, buildDir & "\notes.txt"
    If Err.Number <> 0 Then
        SmallFiles = "Unable to copy " & what & " to the cartridge (" & Err.Description & ")."
        Exit Function
    End If
End Function

' How it's going:
'   packing|percent|files done|files
'   done|anything to mention
'   fail|what went wrong
'   stopped
Public Function EngBuildPoll() As String
    Dim doneFile As String, t As Object, s As String, code As Long, n As Long, pct As Long
    doneFile = buildTmp & "\done.txt"
    ReadLogMore
    s = ""
    If fso.FileExists(doneFile) Then
        On Error Resume Next
        Set t = fso.OpenTextFile(doneFile, 1, False, 0)
        If Err.Number = 0 Then
            If Not t.AtEndOfStream Then s = t.ReadAll
            t.Close
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If InStr(s, "code") = 0 Then
        ' still going
        pct = 0
        If totalBytes > 0 Then
            pct = Int(doneBytes * 100 / totalBytes)
            n = Int(PieceBytes() * 100 / totalBytes)
            If n > pct Then pct = n
        End If
        If pct > 99 Then pct = 99
        EngBuildPoll = "packing|" & pct & "|" & doneFiles & "|" & fileCount
        Exit Function
    End If
    code = CLng(Val2(Mid(s, InStr(s, "code") + 4)))
    If buildStopped Then
        RemovePak
        EngBuildPoll = "stopped"
        Exit Function
    End If
    If code > 1 Then
        EngBuildPoll = "fail|7-Zip exited with error " & code & "." & LogTrouble() & vbLf & "The incomplete cartridge is in " & buildDir & "."
        Exit Function
    End If
    n = CountPieces(buildDir)
    If n = 0 Then
        EngBuildPoll = "fail|7-Zip did not create any archive files." & LogTrouble() & vbLf & "The incomplete cartridge is in " & buildDir & "."
        Exit Function
    End If
    ' and last, the details, so the Cartridge Server only ever sees a finished cartridge
    On Error Resume Next
    WriteIni buildDir & "\pak.ini", buildIconName, buildIconFrom, n
    If Err.Number <> 0 Then
        EngBuildPoll = "fail|Unable to write pak.ini (" & Err.Description & ")."
        Exit Function
    End If
    On Error GoTo 0
    s = ""
    If code = 1 Then s = "7-Zip could not read some files. They may be open in another program:" & LogTrouble()
    EngBuildPoll = "done|" & s
End Function

' What 7-Zip has said since last time: each file it lists is done
Private Sub ReadLogMore()
    Dim logFile As String, t As Object, s As String, lns() As String, i As Long, ln As String, nm As String
    logFile = buildTmp & "\7za.log"
    If Not fso.FileExists(logFile) Then Exit Sub
    On Error Resume Next
    Set t = fso.OpenTextFile(logFile, 1, False, 0)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    s = ""
    If Not t.AtEndOfStream Then s = t.ReadAll
    t.Close
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    ' just the part we haven't seen yet
    If Len(s) <= logPos Then Exit Sub
    s = Mid(s, logPos + 1)
    logPos = logPos + Len(s)
    s = logRest & Replace(s, vbCr, "")
    lns = Split(s, vbLf)
    ' the last bit may be half a line
    logRest = lns(UBound(lns))
    For i = 0 To UBound(lns) - 1
        ln = lns(i)
        nm = ""
        If Left(ln, 13) = "Compressing  " Then nm = Mid(ln, 14)
        If newer7z And Left(ln, 2) = "+ " Then nm = Mid(ln, 3)
        If Len(nm) > 0 Then
            doneFiles = doneFiles + 1
            If buildSizes.Exists(nm) Then doneBytes = doneBytes + buildSizes(nm)
        ElseIf InStr(1, ln, "WARNING", vbTextCompare) > 0 Or InStr(1, ln, "Error", vbTextCompare) > 0 Then
            If Len(buildWarn) < 2000 Then buildWarn = buildWarn & vbLf & Trim(ln)
        End If
    Next
End Sub

' What 7-Zip complained about, or the end of its log
Private Function LogTrouble() As String
    Dim t As Object, s As String, lns() As String, i As Long, out As String, n As Long
    LogTrouble = ""
    If Len(buildWarn) > 0 Then
        LogTrouble = buildWarn
        Exit Function
    End If
    On Error Resume Next
    Set t = fso.OpenTextFile(buildTmp & "\7za.log", 1, False, 0)
    If Err.Number <> 0 Then Exit Function
    s = ""
    If Not t.AtEndOfStream Then s = t.ReadAll
    t.Close
    lns = Split(Replace(s, vbCr, ""), vbLf)
    out = ""
    n = 0
    For i = UBound(lns) To 0 Step -1
        If Len(Trim(lns(i))) > 0 Then
            out = vbLf & Trim(lns(i)) & out
            n = n + 1
            If n = 4 Then Exit For
        End If
    Next
    LogTrouble = out
End Function

' The pieces made so far, altogether
Private Function PieceBytes() As Double
    Dim f As Object, n As Double
    n = 0
    On Error Resume Next
    For Each f In fso.GetFolder(buildDir).Files
        If LCase(Left(f.Name, 8)) = "game.7z." Then n = n + CDbl(f.Size)
    Next
    PieceBytes = n
End Function

' Stop packing. The next EngBuildPoll tidies up and says "stopped".
Public Sub EngBuildStop()
    buildStopped = True
    KillSevenZip
End Sub

' The window's closing in the middle of packing: stop, and take away the half-made cartridge
Public Sub EngBuildAbandon()
    Dim k As Long
    buildStopped = True
    KillSevenZip
    For k = 1 To 5
        If fso.FileExists(buildTmp & "\done.txt") Then Exit For
        Pause
    Next
    RemovePak
End Sub

Private Sub KillSevenZip()
    Dim wmi As Object, p As Object, cl As String
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If Err.Number <> 0 Then Exit Sub
    For Each p In wmi.ExecQuery("Select * From Win32_Process")
        If LCase(p.Name) = "7za.exe" Then
            cl = ""
            cl = p.CommandLine
            If Len(cl) = 0 Or InStr(1, cl, buildDir, vbTextCompare) > 0 Then p.Terminate 1
        End If
    Next
End Sub

Private Sub RemovePak()
    Dim k As Long
    On Error Resume Next
    For k = 1 To 3
        If Not fso.FolderExists(buildDir) Then Exit Sub
        fso.DeleteFolder buildDir, True
        Err.Clear
        If fso.FolderExists(buildDir) Then Pause
    Next
End Sub

' About a second: ping waits for us, out of sight
Private Sub Pause()
    On Error Resume Next
    sh.Run "ping -n 2 127.0.0.1", 0, True
End Sub

' 7-Zip 9.20 vs newer: their first line says which
Private Function SevenZipIsNewer() As Boolean
    Dim verFile As String, t As Object, first As String, p As Long, v As Double
    SevenZipIsNewer = False
    verFile = buildTmp & "\7za-version.txt"
    On Error Resume Next
    sh.Run "cmd /c """"" & here & "\7za.exe"" > """ & verFile & """""", 0, True
    Set t = fso.OpenTextFile(verFile, 1, False, 0)
    If Err.Number <> 0 Then Exit Function
    first = ""
    Do While Not t.AtEndOfStream And Len(first) = 0
        first = Trim(t.ReadLine)
    Loop
    t.Close
    On Error GoTo 0
    p = InStr(first, ")")
    If p > 0 Then
        v = Val2(Mid(first, p + 1))
        SevenZipIsNewer = (v >= 15)
    End If
End Function

' ---------------------------------------------------------------- finding the games in the folder

Private Sub StartCandidates()
    cCount = 0
    ReDim cName(60)
    ReDim cExe(60)
    ReDim cArgs(60)
    ReDim cStart(60)
    ReDim cIcon(60)
    ReDim cPlayers(60)
    ReDim cCfg(60)
    ReDim cCfgArgs(60)
End Sub

Private Sub AddCandidate(ByVal nm As String, ByVal ex As String, ByVal ar As String, ByVal st As String, ByVal ic As String)
    If cCount > 59 Then Exit Sub
    cName(cCount) = nm
    cExe(cCount) = ex
    cArgs(cCount) = Trim(ar)
    cStart(cCount) = st
    cIcon(cCount) = ic
    cPlayers(cCount) = gPlayers
    cCfg(cCount) = ""
    cCfgArgs(cCount) = ""
    cCount = cCount + 1
End Sub

' Shortcuts in the folder first: they say exactly how each game starts. Then any mod
' folders with a liblist.gam (Half-Life and the games built on it) they didn't cover.
Private Sub DetectGames()
    Dim i As Long, lnk As Object, target As String, wd As String, icn As String, p As Long
    Dim f As Object, nm As String, launcher As String, modDir As String, gameArgs As String
    StartCandidates
    On Error Resume Next
    For i = 0 To lnkCount - 1
        Set lnk = sh.CreateShortcut(srcDir & "\" & lnkList(i))
        If Err.Number = 0 Then
            target = InsideGame(lnk.TargetPath)
            If Len(target) > 0 And LCase(fso.GetExtensionName(target)) = "exe" Then
                wd = InsideGame(lnk.WorkingDirectory)
                If LCase(wd) = LCase(fso.GetParentFolderName(target)) Then wd = ""
                icn = lnk.IconLocation
                p = InStrRev(icn, ",")
                If p > 0 Then icn = Left(icn, p - 1)
                icn = InsideGame(icn)
                If LCase(icn) = LCase(target) Then icn = ""
                AddCandidate fso.GetBaseName(lnkList(i)), target, lnk.Arguments, wd, icn
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
    ' shortcuts that all share the same extra settings: those go after every game instead
    SplitPackArgs
    ' Half-Life style mods: a folder with a liblist.gam, started with -game <folder>
    launcher = ""
    If fso.FileExists(srcDir & "\hl.exe") Then launcher = "hl.exe"
    If Len(launcher) = 0 Then launcher = GuessExe()
    For Each f In fso.GetFolder(srcDir).SubFolders
        If fso.FileExists(f.Path & "\liblist.gam") Then
            modDir = f.Name
            If Not ModCovered(modDir) Then
                nm = LiblistName(f.Path & "\liblist.gam")
                If Len(nm) = 0 Then nm = modDir
                If LCase(modDir) = "valve" Then
                    gameArgs = ""
                Else
                    gameArgs = "-game " & modDir
                End If
                AddCandidate nm, launcher, gameArgs, "", ModIcon(f)
            End If
        End If
    Next
End Sub

' A path from a shortcut, as a path inside the game folder. Shortcuts made on another PC
' point somewhere else, so a file with the same name in the game folder counts too.
Private Function InsideGame(ByVal p As String) As String
    Dim r As String
    InsideGame = ""
    If Len(p) = 0 Then Exit Function
    r = RelPath(p)
    If LCase(r) <> LCase(Replace(p, "/", "\")) Then
        If fso.FileExists(srcDir & "\" & r) Or fso.FolderExists(srcDir & "\" & r) Then InsideGame = r
        Exit Function
    End If
    If fso.FileExists(srcDir & "\" & fso.GetFileName(p)) Then InsideGame = fso.GetFileName(p)
End Function

' A shortcut already starts this mod
Private Function ModCovered(ByVal modDir As String) As Boolean
    Dim i As Long
    ModCovered = False
    For i = 0 To cCount - 1
        If InStr(1, " " & cArgs(i) & " ", " -game " & modDir & " ", vbTextCompare) > 0 Then ModCovered = True
        If LCase(modDir) = "valve" And InStr(1, cArgs(i), "-game ", vbTextCompare) = 0 And LCase(fso.GetFileName(cExe(i))) = "hl.exe" Then ModCovered = True
    Next
End Function

' The game's name from a liblist.gam:  game "Counter-Strike"
Private Function LiblistName(ByVal path As String) As String
    Dim t As Object, lns() As String, i As Long, l As String, q1 As Long, q2 As Long
    LiblistName = ""
    On Error Resume Next
    Set t = fso.OpenTextFile(path, 1, False, 0)
    lns = Split(Replace(t.ReadAll, vbCr, ""), vbLf)
    t.Close
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    For i = 0 To UBound(lns)
        l = Trim(Replace(lns(i), vbTab, " "))
        If LCase(Left(l, 5)) = "game " Then
            q1 = InStr(l, """")
            q2 = InStrRev(l, """")
            If q1 > 0 And q2 > q1 Then
                LiblistName = Tidy(Mid(l, q1 + 1, q2 - q1 - 1))
                Exit Function
            End If
        End If
    Next
End Function

' A mod's icon: game.ico (Steam-era mods have one), or any other .ico in its folder
Private Function ModIcon(ByVal folder As Object) As String
    Dim f As Object
    ModIcon = ""
    If fso.FileExists(folder.Path & "\game.ico") Then
        ModIcon = RelPath(folder.Path & "\game.ico")
        Exit Function
    End If
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "ico" Then
            ModIcon = RelPath(f.Path)
            Exit Function
        End If
    Next
End Function

' When every shortcut has the same settings besides its own -game, like
'   hl.exe -game cstrike -nomaster -steam -num_edicts 4096
' the common part is suggested for the whole multicart and each game keeps just -game cstrike
Private Sub SplitPackArgs()
    Dim i As Long, re As Object, own As String, rest As String, common As Variant, same As Boolean
    If cCount < 2 Then Exit Sub
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "-game\s+\S+"
    same = True
    common = Empty
    For i = 0 To cCount - 1
        own = ""
        If re.Test(cArgs(i)) Then own = re.Execute(cArgs(i)).Item(0).Value
        rest = Squash(Replace(cArgs(i), own, ""))
        If IsEmpty(common) Then
            common = rest
        ElseIf rest <> common Then
            same = False
        End If
    Next
    If Not same Or Len(common) = 0 Then Exit Sub
    For i = 0 To cCount - 1
        own = ""
        If re.Test(cArgs(i)) Then own = re.Execute(cArgs(i)).Item(0).Value
        If Len(own) > 0 Or Squash(cArgs(i)) = common Then cArgs(i) = own
    Next
    sharedArgs = common
End Sub

Private Function Squash(ByVal s As String) As String
    Squash = Trim(s)
    Do While InStr(Squash, "  ") > 0
        Squash = Replace(Squash, "  ", " ")
    Loop
End Function

' Each game's icon on a multicart: a file inside the game is left for the Cartridge Server to take
' out of the archive (icon_from). A picture from somewhere else is copied into the cartridge.
Private Sub PackIcons(ByVal pakFolder As String)
    Dim i As Long, ic As String, ext As String, rel As String, inside As Boolean
    ReDim gmIconFile(gmCount)
    ReDim gmIconFrom(gmCount)
    For i = 0 To gmCount - 1
        ic = gmIcon(i)
        gmIconFile(i) = ""
        gmIconFrom(i) = ""
        If Len(ic) = 0 Then
            gmIconFrom(i) = gmExe(i)
        Else
            rel = RelPath(ic)
            ext = LCase(fso.GetExtensionName(ic))
            ' inside the game: either it was given as a path in the game folder, or as a full path into it
            inside = False
            If fso.FileExists(srcDir & "\" & rel) Then
                If LCase(rel) <> LCase(Replace(ic, "/", "\")) Then
                    inside = True
                ElseIf Not fso.FileExists(ic) Then
                    inside = True
                End If
            End If
            If inside Then
                gmIconFrom(i) = rel
            Else
                gmIconFile(i) = "icon-" & gmId(i) & "." & ext
                fso.CopyFile ic, pakFolder & "\" & gmIconFile(i)
            End If
        End If
    Next
End Sub

' How many pieces 7-Zip made, and how big they are altogether (the Cartridge Server and
' the LAN Center use that to tell a finished cartridge from one still being copied)
Private Function CountPieces(ByVal pakFolder As String) As Long
    Dim f As Object, n As Long
    n = 0
    archiveBytes = 0
    For Each f In fso.GetFolder(pakFolder).Files
        If LCase(Left(f.Name, 8)) = "game.7z." Then
            n = n + 1
            archiveBytes = archiveBytes + CDbl(f.Size)
        End If
    Next
    CountPieces = n
End Function

Private Sub WriteIni(ByVal path As String, ByVal iconName As String, ByVal iconFrom As String, ByVal pieces As Long)
    Dim t As Object, i As Long
    Set t = fso.CreateTextFile(path, True, False)
    t.WriteLine "; Palms Arcade game cartridge - made with Palms LAN Center " & makerVer
    t.WriteLine "; To rebuild it: Game Library > Make Cartridge > Rebuild a cartridge, then select this file"
    t.WriteLine "[pak]"
    If isPack Then
        t.WriteLine "format = 2"
    Else
        t.WriteLine "format = 1"
    End If
    t.WriteLine "id = " & gId
    t.WriteLine "name = " & gName
    t.WriteLine "version = " & gVer
    t.WriteLine "revision = " & gRev
    t.WriteLine "players = " & gPlayers
    t.WriteLine "folder = " & gFolder
    If isPack Then
        ' a multicart: each game below starts its own way. args here go after every one of them.
        t.WriteLine "exe = "
        t.WriteLine "args = " & packArgs
        t.WriteLine "config_exe = "
        t.WriteLine "config_args = "
        t.WriteLine "games = " & gmCount
    Else
        t.WriteLine "exe = " & gExe
        t.WriteLine "args = " & gArgs
        t.WriteLine "config_exe = " & gCfg
        t.WriteLine "config_args = " & gCfgArgs
    End If
    t.WriteLine "icon = " & iconName
    t.WriteLine "icon_from = " & iconFrom
    If Len(gReg) > 0 Then
        t.WriteLine "reg = setup.reg"
    Else
        t.WriteLine "reg = "
    End If
    If Len(gNotes) > 0 Then
        t.WriteLine "notes = notes.txt"
    Else
        t.WriteLine "notes = "
    End If
    t.WriteLine "archive = game.7z"
    t.WriteLine "pieces = " & pieces
    t.WriteLine "archive_bytes = " & FormatNumber(archiveBytes, 0, , , False)
    t.WriteLine "files = " & fileCount
    t.WriteLine "size = " & FormatNumber(totalBytes, 0, , , False)
    t.WriteLine "made = " & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2)
    t.WriteLine ""
    If isPack Then
        For i = 0 To gmCount - 1
            t.WriteLine "[game" & (i + 1) & "]"
            t.WriteLine "id = " & gmId(i)
            t.WriteLine "name = " & gmName(i)
            t.WriteLine "players = " & gmPlayers(i)
            t.WriteLine "exe = " & gmExe(i)
            t.WriteLine "args = " & gmArgs(i)
            t.WriteLine "start_in = " & gmStart(i)
            t.WriteLine "config_exe = " & gmCfg(i)
            t.WriteLine "config_args = " & gmCfgArgs(i)
            t.WriteLine "icon = " & gmIconFile(i)
            t.WriteLine "icon_from = " & gmIconFrom(i)
            t.WriteLine "icon_source = " & gmIcon(i)
            t.WriteLine ""
        Next
    End If
    t.WriteLine "; Source settings, used when rebuilding"
    t.WriteLine "[maker]"
    t.WriteLine "version = " & makerVer
    t.WriteLine "source = " & srcDir
    t.WriteLine "icon_source = " & gIcon
    t.WriteLine "reg_source = " & gReg
    t.WriteLine "reg_path = " & gRegPath
    t.WriteLine "notes_source = " & gNotes
    t.WriteLine "output = " & outRoot
    t.Close
End Sub

' The registry file, ready for any PC:
'  - in the old REGEDIT4 format, which Windows 98 and XP both understand
'  - the folder it was installed in swapped for %INSTALLDIR_REG% (and %INSTALLDIR%),
'    which the LAN Center fills in with the folder each PC installs to
Private Sub WriteReg(ByVal src As String, ByVal dst As String, ByVal oldPath As String)
    Dim text As String
    text = CleanReg(ReadAnyText(src))
    If Len(oldPath) > 0 Then
        text = Replace(text, Replace(oldPath, "\", "\\"), "%INSTALLDIR_REG%", 1, -1, vbTextCompare)
        text = Replace(text, oldPath, "%INSTALLDIR%", 1, -1, vbTextCompare)
    End If
    WritePlain dst, text
End Sub

' Plain text (ANSI), which every Windows reads. The odd character that won't fit becomes ?
Private Sub WritePlain(ByVal path As String, ByVal text As String)
    Dim t As Object, i As Long, ch As String, safe As String
    On Error Resume Next
    Set t = fso.CreateTextFile(path, True, False)
    t.Write text
    t.Close
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    safe = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If AscW(ch) < 0 Or AscW(ch) > 255 Then ch = "?"
        safe = safe & ch
    Next
    Set t = fso.CreateTextFile(path, True, False)
    t.Write safe
    t.Close
    On Error GoTo 0
End Sub

' Windows 98 exports registry files as plain text starting REGEDIT4. XP and Vista export
' "Windows Registry Editor Version 5.00" files in Unicode, where two kinds of setting
' (hex(2) and hex(7)) hold their text as Unicode too. Only REGEDIT4 works on all of them,
' and a file without its first line is ignored by regedit altogether. So whatever comes
' in, a clean REGEDIT4 file goes in the cartridge, and anything odd is pointed out.
' Sets regKeys, regValues, regNotes (things fixed) and regTrouble (things that look wrong).
Private Function CleanReg(ByVal text As String) As String
    Dim lns() As String, i As Long, h As Long, k As Long, head As String, ver As Long, body As String, t As String
    Dim logical As String, first As Long, re As Object, m As Object, keyOpen As Boolean, wowDone As Boolean
    Dim shortDone As Boolean, lost As Boolean, troubles As Long
    regKeys = 0
    regValues = 0
    regNotes = ""
    regTrouble = ""
    troubles = 0
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    ' the marker some editors put first: a character or three that aren't plain text
    For i = 1 To 3
        If Len(text) = 0 Then Exit For
        If Asc(Left(text, 1)) < 128 And Left(text, 1) <> "?" Then Exit For
        text = Mid(text, 2)
    Next
    lns = Split(text, vbLf)
    ' the first line says what kind of file it is
    h = 0
    Do While h <= UBound(lns)
        If Len(Trim(lns(h))) > 0 Then Exit Do
        h = h + 1
    Loop
    ver = 4
    If h > UBound(lns) Then
        CleanReg = "REGEDIT4" & vbCrLf & vbCrLf
        regTrouble = regTrouble & vbCrLf & "      The file is empty."
        Exit Function
    End If
    head = UCase(Trim(lns(h)))
    If head = "REGEDIT4" Then
        h = h + 1
    ElseIf Left(head, 33) = "WINDOWS REGISTRY EDITOR VERSION 5" Then
        ver = 5
        h = h + 1
        regNotes = regNotes & vbCrLf & "    Windows XP/Vista format (Version 5.00). Converted to REGEDIT4 for Windows 98 compatibility."
    ElseIf Left(head, 1) = "[" Then
        regNotes = regNotes & vbCrLf & "    The REGEDIT4 header was missing (regedit ignores files without one). The header was added."
    Else
        regTrouble = regTrouble & vbCrLf & "      Line 1 should be REGEDIT4, but is: " & Left(Trim(lns(h)), 60)
        troubles = troubles + 1
        h = h + 1
    End If
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    body = ""
    keyOpen = False
    i = h
    Do While i <= UBound(lns)
        ' a setting can carry on over several lines, each but the last ending in \
        first = i
        logical = RTrim(lns(i))
        Do While Right(logical, 1) = "\" And i < UBound(lns)
            i = i + 1
            logical = Left(logical, Len(logical) - 1) & Trim(lns(i))
        Loop
        i = i + 1
        t = Trim(logical)
        If Len(t) = 0 Or Left(t, 1) = ";" Then
            body = body & t & vbCrLf
        ElseIf Left(t, 1) = "[" Then
            ' a key: HKLM and friends spelled out in full, which regedit needs
            re.Pattern = "^\[(-?)HK(LM|CU|CR|U|CC)(\\|\])"
            If re.Test(t) Then
                Set m = re.Execute(t)(0)
                t = "[" & m.SubMatches(0) & LongRoot(m.SubMatches(1)) & Mid(t, Len(m.Value))
                If Not shortDone Then regNotes = regNotes & vbCrLf & "    Abbreviated root keys (such as HKLM) were expanded to their full names."
                shortDone = True
            End If
            ' exported on 64-bit Windows: put it back where a 32-bit game looks
            If InStr(1, t, "\SOFTWARE\Wow6432Node\", vbTextCompare) > 0 Then
                t = Replace(t, "\SOFTWARE\Wow6432Node\", "\SOFTWARE\", 1, 1, vbTextCompare)
                If Not wowDone Then regNotes = regNotes & vbCrLf & "    Keys under Wow6432Node (exported on 64-bit Windows) were moved to the standard 32-bit location."
                wowDone = True
            End If
            re.Pattern = "^\[-?HKEY_(LOCAL_MACHINE|CURRENT_USER|CLASSES_ROOT|USERS|CURRENT_CONFIG|DYN_DATA)(\\[^\]]*)?\]$"
            If re.Test(t) Then
                If Left(t, 2) <> "[-" Then regKeys = regKeys + 1
                keyOpen = (Left(t, 2) <> "[-")
            ElseIf troubles < 4 Then
                regTrouble = regTrouble & vbCrLf & "      Line " & (first + 1) & ": unknown root key: " & Left(t, 60)
                troubles = troubles + 1
            End If
            body = body & t & vbCrLf
        ElseIf Left(t, 1) = """" Or Left(t, 1) = "@" Then
            If Not keyOpen And troubles < 4 Then
                regTrouble = regTrouble & vbCrLf & "      Line " & (first + 1) & ": value without a [key] section."
                troubles = troubles + 1
            End If
            regValues = regValues + 1
            re.Pattern = "^(""(?:[^""\\]|\\.)*""|@)\s*=\s*hex\((2|7)\):(.*)$"
            If ver = 5 And re.Test(t) Then
                Set m = re.Execute(t)(0)
                body = body & HexUnicodeToPlain(m.SubMatches(0), m.SubMatches(1), m.SubMatches(2), lost)
            Else
                ' as it was, line breaks and all
                For k = first To i - 1
                    body = body & RTrim(lns(k)) & vbCrLf
                Next
            End If
        ElseIf troubles < 4 Then
            regTrouble = regTrouble & vbCrLf & "      Line " & (first + 1) & ": not a valid registry line: " & Left(t, 60)
            troubles = troubles + 1
        End If
    Loop
    If lost Then regNotes = regNotes & vbCrLf & "    Some Unicode characters could not be converted and were replaced with ? (not supported by Windows 98)."
    If regValues = 0 And troubles < 4 Then regTrouble = regTrouble & vbCrLf & "      The file contains no values."
    ' tidy blank lines at the start, and end with a blank line, which old regedit needs
    Do While Left(body, 2) = vbCrLf
        body = Mid(body, 3)
    Loop
    Do While Right(body, 4) = vbCrLf & vbCrLf
        body = Left(body, Len(body) - 2)
    Loop
    CleanReg = "REGEDIT4" & vbCrLf & vbCrLf & body & vbCrLf
End Function

Private Function LongRoot(ByVal s As String) As String
    Select Case UCase(s)
    Case "LM": LongRoot = "HKEY_LOCAL_MACHINE"
    Case "CU": LongRoot = "HKEY_CURRENT_USER"
    Case "CR": LongRoot = "HKEY_CLASSES_ROOT"
    Case "U": LongRoot = "HKEY_USERS"
    Case Else: LongRoot = "HKEY_CURRENT_CONFIG"
    End Select
End Function

' A hex(2) or hex(7) setting from a Version 5.00 file holds Unicode: two bytes a letter.
' REGEDIT4 wants one byte a letter. Returns the setting's line(s), ready for the file.
Private Function HexUnicodeToPlain(ByVal nm As String, ByVal kind As String, ByVal data As String, lost As Boolean) As String
    Dim bytes() As String, i As Long, lo As Long, hi As Long, out As String, n As Long
    bytes = Split(Replace(Replace(data, " ", ""), vbTab, ""), ",")
    out = ""
    n = 0
    For i = 0 To UBound(bytes) - 1 Step 2
        lo = CLng("&H" & bytes(i))
        hi = CLng("&H" & bytes(i + 1))
        If hi <> 0 Then
            lo = 63
            lost = True
        End If
        If n > 0 Then out = out & ","
        ' a line break every 20 bytes, the way regedit writes them
        If n > 0 And n Mod 20 = 0 Then out = out & "\" & vbCrLf & "  "
        out = out & Right("0" & LCase(Hex(lo)), 2)
        n = n + 1
    Next
    HexUnicodeToPlain = nm & "=hex(" & kind & "):" & out & vbCrLf
End Function

' Reads a text file whether it's plain or Unicode (XP's regedit saves Unicode)
Private Function ReadAnyText(ByVal path As String) As String
    Dim t As Object, head As String, uni As Boolean, st As Object
    uni = False
    If fso.GetFile(path).Size >= 3 Then
        Set t = fso.OpenTextFile(path, 1, False, 0)
        head = t.Read(3)
        t.Close
        ' UTF-8 files (newer Notepad) start with the bytes EF BB BF
        If Asc(Left(head, 1)) = 239 And Asc(Mid(head, 2, 1)) = 187 And Asc(Mid(head, 3, 1)) = 191 Then
            On Error Resume Next
            Set st = CreateObject("ADODB.Stream")
            st.Type = 2
            st.Charset = "utf-8"
            st.Open
            st.LoadFromFile path
            ReadAnyText = st.ReadText
            st.Close
            If Err.Number = 0 Then Exit Function
            Err.Clear
            On Error GoTo 0
            ' no ADODB: plain text without those three bytes is close enough for a registry file
            Set t = fso.OpenTextFile(path, 1, False, 0)
            ReadAnyText = Mid(t.ReadAll, 4)
            t.Close
            Exit Function
        End If
    End If
    If fso.GetFile(path).Size >= 2 Then
        Set t = fso.OpenTextFile(path, 1, False, 0)
        head = t.Read(2)
        t.Close
        ' Unicode files start with the bytes FF FE
        uni = (Asc(Left(head, 1)) = 255 And Asc(Mid(head, 2, 1)) = 254)
    End If
    If uni Then
        Set t = fso.OpenTextFile(path, 1, False, -1)
    Else
        Set t = fso.OpenTextFile(path, 1, False, 0)
    End If
    If t.AtEndOfStream Then
        ReadAnyText = ""
    Else
        ReadAnyText = t.ReadAll
    End If
    t.Close
End Function

' The first folder that looks like where the game was installed, like C:\SIERRA\Half-Life
Private Function PickRegPath(ByVal text As String) As String
    Dim re As Object, m As Object, p As String
    PickRegPath = ""
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """([A-Za-z]:\\\\[^""]*)"""
    re.Global = True
    For Each m In re.Execute(text)
        p = Replace(m.SubMatches(0), "\\", "\")
        If Right(p, 1) = "\" And Len(p) > 3 Then p = Left(p, Len(p) - 1)
        ' a file (C:\Games\HL\hl.exe) means its folder
        If Len(fso.GetExtensionName(p)) > 0 And Len(fso.GetExtensionName(p)) <= 4 Then p = fso.GetParentFolderName(p)
        If Len(p) > 3 Then
            PickRegPath = p
            Exit Function
        End If
    Next
End Function

Private Function GuessExe() As String
    Dim i As Long, best As String, nm As String, key As String
    GuessExe = ""
    If exeCount = 0 Then Exit Function
    best = exeList(0)
    key = LCase(Replace(gName, " ", ""))
    For i = 0 To exeCount - 1
        nm = LCase(fso.GetBaseName(exeList(i)))
        If InStr(key, nm) > 0 Or InStr(nm, key) > 0 Or nm = gId Then
            GuessExe = exeList(i)
            Exit Function
        End If
        ' skip the usual helpers
        If InStr(" setup uninst unwise uninstall ucc config readme dxsetup ", " " & nm & " ") > 0 Then
            If LCase(best) = LCase(exeList(i)) And i + 1 < exeCount Then best = exeList(i + 1)
        End If
    Next
    GuessExe = best
End Function

' A path inside the game folder, without the game folder on the front
Private Function RelPath(ByVal p As String) As String
    Dim a As String
    a = Replace(p, "/", "\")
    If Len(srcDir) > 0 Then
        If LCase(Left(a, Len(srcDir) + 1)) = LCase(srcDir & "\") Then a = Mid(a, Len(srcDir) + 2)
    End If
    Do While Left(a, 1) = "\"
        a = Mid(a, 2)
    Loop
    RelPath = a
End Function

Private Function MakeId(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    r = ""
    For i = 1 To Len(s)
        c = LCase(Mid(s, i, 1))
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then r = r & c
    Next
    If Len(r) > 24 Then r = Left(r, 24)
    If Len(r) < 2 Then r = "game" & r
    MakeId = r
End Function

Private Function IsId(ByVal s As String) As Boolean
    Dim i As Long, c As String
    IsId = False
    If Len(s) < 2 Or Len(s) > 24 Then Exit Function
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If Not ((c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = "-") Then Exit Function
    Next
    IsId = True
End Function

Private Function IsPlayers(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]{1,3}(-[0-9]{1,3})?$"
    IsPlayers = re.Test(s)
End Function

' Letters, numbers, - _ and . only: safe as a folder name for any old game
Private Function MakeFolderName(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    r = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Or c = "-" Or c = "_" Or c = "." Then r = r & c
    Next
    If Len(r) = 0 Then r = "Game"
    MakeFolderName = r
End Function

' Takes out characters that would upset the pak.ini file
Private Function Tidy(ByVal s As String) As String
    Tidy = Trim(Replace(Replace(Replace(s, "|", "-"), "~", "-"), ";", ","))
End Function

Private Function MakeFolder(ByVal p As String) As Boolean
    Dim parent As String
    MakeFolder = False
    If Len(p) = 0 Then Exit Function
    If fso.FolderExists(p) Then
        MakeFolder = True
        Exit Function
    End If
    parent = fso.GetParentFolderName(p)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        If Not MakeFolder(parent) Then Exit Function
    End If
    On Error Resume Next
    fso.CreateFolder p
    On Error GoTo 0
    MakeFolder = fso.FolderExists(p)
End Function

Private Function SizeText2(ByVal n As Double) As String
    If n >= 1073741824 Then
        SizeText2 = FormatNumber(n / 1073741824, 1) & " GB"
    ElseIf n >= 1048576 Then
        SizeText2 = FormatNumber(n / 1048576, 0) & " MB"
    Else
        SizeText2 = FormatNumber(n / 1024, 0) & " KB"
    End If
End Function

' The number at the start of a bit of text, like 9.20 from " 9.20  Copyright..."
Private Function Val2(ByVal s As String) As Double
    Dim re As Object, m As Object
    Val2 = 0
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[0-9]+(\.[0-9]+)?"
    Set m = re.Execute(s)
    If m.Count > 0 Then Val2 = CDbl(Replace(m(0).Value, ".", Mid(CStr(1.5), 2, 1)))
End Function

' The answers from a cartridge made before, so rebuilding is mostly clicking Next
Private Sub LoadPrevious(ByVal path As String)
    Dim t As Object, lns() As String, i As Long, ln As String, p As Long, k As String, v As String, sect As String, g As Object
    prevPath = path
    Set t = fso.OpenTextFile(path, 1, False, 0)
    lns = Split(Replace(t.ReadAll, vbCr, ""), vbLf)
    t.Close
    sect = ""
    For i = 0 To UBound(lns)
        ln = Trim(lns(i))
        If Left(ln, 1) = "[" And Right(ln, 1) = "]" Then
            sect = LCase(Mid(ln, 2, Len(ln) - 2))
            If Left(sect, 4) = "game" And IsNumeric(Mid(sect, 5)) Then
                ' one of a multicart's games
                prevGameCount = prevGameCount + 1
                Set g = CreateObject("Scripting.Dictionary")
                g.CompareMode = 1
                prevGames.Add prevGameCount, g
            End If
        Else
            p = InStr(ln, "=")
            If p > 1 And Left(ln, 1) <> ";" Then
                k = LCase(Trim(Left(ln, p - 1)))
                v = Trim(Mid(ln, p + 1))
                If Left(sect, 4) = "game" And IsNumeric(Mid(sect, 5)) Then
                    Set g = prevGames(prevGameCount)
                    g(k) = v
                ElseIf Not prev.Exists(k) Then
                    ' [maker] comes after [pak]: its "version" is the maker's, so don't let it win
                    prev(k) = v
                End If
            End If
        End If
    Next
End Sub

Private Function GetKey(ByVal d As Object, ByVal k As String) As String
    GetKey = ""
    If d.Exists(k) Then GetKey = d(k)
End Function

Private Function PakDir() As String
    ' quake-v1.09-r2: a new revision of the same version gets its own folder
    If Len(gVer) > 0 Then
        PakDir = outRoot & "\" & gId & "-" & MakeFolderName(gVer) & "-r" & gRev
    Else
        PakDir = outRoot & "\" & gId & "-r" & gRev
    End If
End Function
