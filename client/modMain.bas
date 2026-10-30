Attribute VB_Name = "modMain"
' Palms LAN Center - Start-up and shared helpers used everywhere.
Option Explicit

' Settings. Changed in Tools > Options and Tools > Notifications. Saved in the registry.
Public gSoundOn As Boolean          ' play sounds at all
Public gNotifyPA As Boolean         ' pop up announcements from the host
Public gNotifyVote As Boolean       ' pop up when a vote starts
Public gNotifySoon As Boolean       ' pop up when voting is about to close
Public gNotifyResult As Boolean     ' pop up the vote result
Public gNotifyName As Boolean       ' ding and pop up when someone says my name
Public gNotifyNext As Boolean       ' pop up 15, 10 and 5 minutes before the next thing on the schedule
Public gNotifyCups As Boolean       ' pop up tournament sign-ups, your match being ready, and the winner
Public gVolume As Long              ' 0 to 100, how loud our own sounds play
Public gClassicPopups As Boolean    ' True = a box in the middle that waits for OK, False = Messenger style
Public gInterrupt As Boolean        ' pop-ups may jump in front of full-screen games
Public gShowJoins As Boolean        ' show "X has signed in", "X is away" and so on in Chat
Public gIgnore As String            ' nicknames whose chat is hidden on this PC, lower case, separated by ;

' Which of the four pop-up spots in the corner are in use
Public gToastSlot(0 To 3) As Boolean

Private Type GUIDType
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUIDType) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function sndPlayMem Lib "winmm.dll" Alias "sndPlaySoundA" (lpSound As Any, ByVal uFlags As Long) As Long
Private Declare Function sndStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpSound As Long, ByVal uFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_HWNDPARENT As Long = -8
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_NOACTIVATE As Long = &H8000000

Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_MEMORY As Long = &H4

' Each sound, loaded once and turned down to the volume setting.
' A sound plays straight from this memory, so it must stay put while it plays.
Private Type WavBuf
    Name As String
    Vol As Long
    Data() As Byte
End Type
Private mWav(0 To 7) As WavBuf
Private mWavCount As Long
Private Const HWND_TOPMOST As Long = -1

' The Winsock control, when Palms LAN Center runs from the zip (Setup registers it otherwise).
' Windows looks for MSWINSCK.OCX in the program's folder first.
Private Declare Function WinsockRegister Lib "MSWINSCK.OCX" Alias "DllRegisterServer" () As Long
Private Declare Function OcxOpenKey Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function OcxQueryValue Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function OcxCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const OCX_KEY_READ As Long = &H20019
Private Const WINSOCK_CLSID As String = "CLSID\{248DD890-BB45-11CF-9ABC-0080C7E7B78D}\InprocServer32"
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Public Sub Main()
    Dim cmdLine As String, id As String
    ' settings from palms.ini and gamepaks.ini (older versions): read them in, once
    ImportOldFiles
    ' Add/Remove Programs starts PalmsLAN.exe /uninstall <cartridge>
    cmdLine = LCase$(Trim$(Command$))
    If Left$(cmdLine, 10) = "/uninstall" Then
        id = Trim$(Mid$(cmdLine, 11))
        If App.PrevInstance Then
            MsgBox "Palms LAN Center is already open." & vbCrLf & vbCrLf & _
                   "To uninstall " & RegGet(GamesKey(id), "name", id) & ", select it in the Game Library and click Uninstall.", vbInformation, "Palms LAN Center"
            Exit Sub
        End If
    End If
    If Not WinsockReady() Then Exit Sub
    LoadSettings
    frmMain.Show
    If Len(id) > 0 Then frmMain.UninstallFromCommand id
End Sub

' True if the Winsock control is registered, registering the copy beside PalmsLAN.exe if need be
Private Function WinsockReady() As Boolean
    Dim h As Long, buf As String, n As Long, f As String
    If OcxOpenKey(HKEY_CLASSES_ROOT, WINSOCK_CLSID, 0, OCX_KEY_READ, h) = 0 Then
        buf = String$(512, 0): n = 512
        If OcxQueryValue(h, vbNullString, 0, 0, buf, n) = 0 And n > 1 Then f = Left$(buf, n - 1)
        OcxCloseKey h
        If Len(f) > 0 Then
            If FileHere(f) Then WinsockReady = True: Exit Function
        End If
    End If
    If Not FileHere(App.Path & "\MSWINSCK.OCX") Then
        MsgBox "The Microsoft Winsock Control (MSWINSCK.OCX) is not installed." & vbCrLf & vbCrLf & _
               "Run Palms LAN Center Setup, or copy MSWINSCK.OCX to " & App.Path & ".", vbCritical, "Palms LAN Center"
        Exit Function
    End If
    On Error GoTo Failed
    If WinsockRegister() = 0 Then WinsockReady = True: Exit Function
Failed:
    MsgBox "Unable to register MSWINSCK.OCX." & vbCrLf & vbCrLf & _
           "Log on as an administrator and start Palms LAN Center again, or run Palms LAN Center Setup.", vbCritical, "Palms LAN Center"
End Function

Public Sub LoadSettings()
    gSoundOn = (SettingGet("sounds", "1") <> "0")
    gNotifyName = (SettingGet("alerts", "1") <> "0")
    gNotifyPA = (SettingGet("notify_pa", "1") <> "0")
    gNotifyVote = (SettingGet("notify_vote", "1") <> "0")
    gNotifySoon = (SettingGet("notify_soon", "1") <> "0")
    gNotifyResult = (SettingGet("notify_result", "1") <> "0")
    gNotifyNext = (SettingGet("notify_next", "1") <> "0")
    gNotifyCups = (SettingGet("notify_cups", "1") <> "0")
    gVolume = Val(SettingGet("volume", "80"))
    If gVolume < 0 Then gVolume = 0
    If gVolume > 100 Then gVolume = 100
    gClassicPopups = (SettingGet("popup_style", "messenger") = "classic")
    gInterrupt = (SettingGet("interrupt", "0") = "1")
    gShowJoins = (SettingGet("show_joins", "1") <> "0")
    gIgnore = SettingGet("ignore", "")
End Sub

Public Sub SaveSettings()
    SettingSet "sounds", OneZero(gSoundOn)
    SettingSet "alerts", OneZero(gNotifyName)
    SettingSet "notify_pa", OneZero(gNotifyPA)
    SettingSet "notify_vote", OneZero(gNotifyVote)
    SettingSet "notify_soon", OneZero(gNotifySoon)
    SettingSet "notify_result", OneZero(gNotifyResult)
    SettingSet "notify_next", OneZero(gNotifyNext)
    SettingSet "notify_cups", OneZero(gNotifyCups)
    SettingSet "volume", CStr(gVolume)
    If gClassicPopups Then
        SettingSet "popup_style", "classic"
    Else
        SettingSet "popup_style", "messenger"
    End If
    SettingSet "interrupt", OneZero(gInterrupt)
    SettingSet "show_joins", OneZero(gShowJoins)
    SettingSet "ignore", gIgnore
End Sub

' Is this person's chat hidden on this PC? (Chat > Ignore)
Public Function IsIgnored(ByVal nick As String) As Boolean
    If Len(nick) = 0 Or Len(gIgnore) = 0 Then Exit Function
    IsIgnored = (InStr(1, ";" & gIgnore & ";", ";" & LCase$(nick) & ";", vbBinaryCompare) > 0)
End Function

' Event details travel as one line, with ~ between lines (SCHEDNOTE, SCHEDINFO)
Public Function NoteToText(ByVal s As String) As String
    Dim p() As String, k As Long, t As String
    If Len(s) = 0 Then Exit Function
    p = Split(s, "~")
    For k = 0 To UBound(p)
        If k > 0 Then t = t & vbCrLf
        t = t & Trim$(p(k))
    Next
    NoteToText = t
End Function

Public Function NoteFromText(ByVal s As String) As String
    s = Replace(s, vbCrLf, "~")
    s = Replace(s, vbCr, "~")
    s = Replace(s, vbLf, "~")
    s = Replace(s, "|", "/")
    Do While InStr(s, "~~") > 0
        s = Replace(s, "~~", "~")
    Loop
    s = Trim$(s)
    Do While Left$(s, 1) = "~"
        s = Trim$(Mid$(s, 2))
    Loop
    Do While Right$(s, 1) = "~"
        s = Trim$(Left$(s, Len(s) - 1))
    Loop
    NoteFromText = s
End Function

Private Function OneZero(ByVal b As Boolean) As String
    If b Then
        OneZero = "1"
    Else
        OneZero = "0"
    End If
End Function

' This PC's network name, like PALMS-DEV
Public Function ThisPCName() As String
    Dim buf As String, n As Long
    buf = String$(64, vbNullChar)
    n = 64
    If GetComputerName(buf, n) <> 0 Then ThisPCName = Left$(buf, n)
End Function

' The folder the program is in, always ending with a backslash
Public Function AppDir() As String
    AppDir = App.Path
    If Right$(AppDir, 1) <> "\" Then AppDir = AppDir & "\"
End Function

' A setting, from HKEY_LOCAL_MACHINE\Software\Palms Arcade\LAN Center
Public Function SettingGet(ByVal key As String, ByVal def As String) As String
    SettingGet = RegGet(REG_LAN, key, def)
End Function

Public Sub SettingSet(ByVal key As String, ByVal value As String)
    RegPut REG_LAN, key, value
End Sub

' A cartridge's version for showing: 1.09 (Rev. 2). Either part may be blank.
Public Function VerText(ByVal ver As String, ByVal rev As String) As String
    If Len(rev) = 0 Then
        VerText = ver
    ElseIf Len(ver) = 0 Then
        VerText = "Rev. " & rev
    Else
        VerText = ver & " (Rev. " & rev & ")"
    End If
End Function

' Windows 95, 98 or ME (no cmd.exe, rd /s or netsh there)
Public Function IsWin9x() As Boolean
    IsWin9x = (GetVersion() < 0)       ' the top bit is set on 95, 98 and ME
End Function

' A random ID for this PC, so the server can remember who sat here
Public Function NewId() As String
    Dim g As GUIDType, i As Integer, s As String
    CoCreateGuid g
    s = Right$("00000000" & Hex$(g.Data1), 8) & Right$("0000" & Hex$(g.Data2), 4) & Right$("0000" & Hex$(g.Data3), 4)
    For i = 0 To 7
        s = s & Right$("0" & Hex$(g.Data4(i)), 2)
    Next
    NewId = s
End Function

' Plays dingdong.wav, triple.wav, horn.wav or ding.wav from the program folder,
' at the volume chosen in Settings. Only our sound is turned down, not the whole PC.
Public Sub PlayChime(ByVal chime As String)
    Dim i As Long, slot As Long, f As String
    If Not gSoundOn Or gVolume <= 0 Then Exit Sub
    If chime = "none" Or Len(chime) = 0 Then Exit Sub
    f = AppDir() & chime & ".wav"
    If Len(Dir$(f)) = 0 Then
        Beep
        Exit Sub
    End If
    slot = -1
    For i = 0 To mWavCount - 1
        If mWav(i).Name = chime Then slot = i
    Next
    If slot = -1 Then
        If mWavCount > 7 Then
            sndPlaySound f, SND_ASYNC Or SND_NODEFAULT
            Exit Sub
        End If
        slot = mWavCount
        mWavCount = mWavCount + 1
        mWav(slot).Name = chime
        mWav(slot).Vol = -1
    End If
    If mWav(slot).Vol <> gVolume Then
        sndStop 0, 0                     ' stop anything playing before its memory changes
        mWav(slot).Vol = -1
        If Not LoadScaled(f, gVolume, mWav(slot).Data) Then
            sndPlaySound f, SND_ASYNC Or SND_NODEFAULT
            Exit Sub
        End If
        mWav(slot).Vol = gVolume
    End If
    sndPlayMem mWav(slot).Data(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

' Reads a WAV file and turns it down to vol percent. Works on 16-bit sound,
' which is what our chimes are. Anything else returns False and plays at full volume.
Private Function LoadScaled(ByVal path As String, ByVal vol As Long, b() As Byte) As Boolean
    Dim fnum As Integer, n As Long, p As Long, chunkSize As Long, fmtTag As Integer, fmtBits As Integer
    Dim dataPos As Long, dataLen As Long, s() As Integer, i As Long, cnt As Long
    On Error GoTo Failed
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    n = LOF(fnum)
    If n < 44 Then
        Close #fnum
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #fnum, , b
    Close #fnum
    If ChunkId(b, 0) <> "RIFF" Or ChunkId(b, 8) <> "WAVE" Then Exit Function
    p = 12
    Do While p + 8 <= n
        CopyMemory chunkSize, b(p + 4), 4
        If ChunkId(b, p) = "fmt " Then
            CopyMemory fmtTag, b(p + 8), 2
            CopyMemory fmtBits, b(p + 22), 2
        ElseIf ChunkId(b, p) = "data" Then
            dataPos = p + 8
            dataLen = chunkSize
            Exit Do
        End If
        If chunkSize < 0 Then Exit Function
        p = p + 8 + chunkSize + (chunkSize And 1)
    Loop
    If fmtTag <> 1 Or fmtBits <> 16 Or dataPos = 0 Then Exit Function
    If dataPos + dataLen > n Then dataLen = n - dataPos
    cnt = dataLen \ 2
    If cnt > 0 And vol < 100 Then
        ReDim s(0 To cnt - 1)
        CopyMemory s(0), b(dataPos), cnt * 2
        For i = 0 To cnt - 1
            s(i) = (CLng(s(i)) * vol) \ 100
        Next
        CopyMemory b(dataPos), s(0), cnt * 2
    End If
    LoadScaled = True
    Exit Function
Failed:
    On Error Resume Next
    Close #fnum
End Function

Private Function ChunkId(b() As Byte, ByVal p As Long) As String
    ChunkId = Chr$(b(p)) & Chr$(b(p + 1)) & Chr$(b(p + 2)) & Chr$(b(p + 3))
End Function

' Brings a window to the very front, even over a full-screen game.
' Windows normally refuses this to programs in the background, so we borrow
' the front program's place in the queue for a moment.
Public Sub ForceForeground(ByVal hwnd As Long)
    Dim fg As Long, pid As Long, theirs As Long, mine As Long
    fg = GetForegroundWindow()
    theirs = GetWindowThreadProcessId(fg, pid)
    mine = GetCurrentThreadId()
    If theirs <> mine And theirs <> 0 Then AttachThreadInput mine, theirs, 1
    SetForegroundWindow hwnd
    BringWindowToTop hwnd
    If theirs <> mine And theirs <> 0 Then AttachThreadInput mine, theirs, 0
End Sub

' True if a form is open, without opening it by asking
Public Function IsFormLoaded(ByVal formName As String) As Boolean
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        If StrComp(Forms(i).Name, formName, vbTextCompare) = 0 Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
End Function

' Cuts a pop-up loose from the LAN Center's other windows. Visual Basic ties all of a
' program's windows together, so bringing a pop-up to the front brought the main window
' with it, right over the top of your game. Call it before the window is shown.
' toolWin: no taskbar button. noActivate: clicking it doesn't take the keyboard from your game.
Public Sub LoneWindow(ByVal hwnd As Long, ByVal toolWin As Boolean, ByVal noActivate As Boolean)
    Dim ex As Long
    SetWindowLong hwnd, GWL_HWNDPARENT, 0
    ex = GetWindowLong(hwnd, GWL_EXSTYLE)
    If toolWin Then ex = ex Or WS_EX_TOOLWINDOW
    If noActivate Then ex = ex Or WS_EX_NOACTIVATE
    SetWindowLong hwnd, GWL_EXSTYLE, ex
End Sub

Public Sub MakeTopmost(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub FlashTitle(ByVal hwnd As Long)
    FlashWindow hwnd, 1
End Sub
