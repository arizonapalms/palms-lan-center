VERSION 5.00
Begin VB.Form frmProjector
   BackColor       =   &H0016140B&
   BorderStyle     =   0  'None
   Caption         =   "Palms LAN Center - Projector"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmProjector"
   ShowInTaskbar   =   -1  'True
   Begin VB.Timer tmrSlide
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picSlide
      AutoRedraw      =   -1  'True
      BackColor       =   &H0016140B&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmProjector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - View > Projector: full screen slides for the big screen, at any resolution.
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_SHOWWINDOW As Long = &H40

' colours (Visual Basic writes them blue, green, red)
Private Const P_BG As Long = &H16140B          ' #0B1416
Private Const P_BAR As Long = &H6B6131         ' #31616B teal
Private Const P_ORANGE As Long = &H218EF7      ' #F78E21
Private Const P_YELLOW As Long = &H29DBFF      ' #FFDB29
Private Const P_WHITE As Long = &HFFFFFF
Private Const P_PALE As Long = &HDAD6BF        ' #BFD6DA
Private Const P_LABEL As Long = &HB5AA8C       ' #8CAAB5
Private Const P_DIM As Long = &H928A6B         ' #6B8A92

Private Const VOTE_SHOW As Long = 600          ' a closed vote stays on this long (seconds)
Private Const CUP_SHOW As Long = 3600          ' and a finished tournament

Private mSlide As Long          ' slides shown so far: which one comes next
Private mLeft As Long           ' seconds until the next slide
Private mKind As String         ' the slide on screen: now, sched, cup, vote or welcome
Private mCupId As String        ' its tournament
Private mDirty As Boolean       ' the data has changed: draw again on the next tick
Private mS As Single            ' the scale: 1 = 800 x 600
Private mW As Long              ' the screen, in pixels
Private mH As Long
Private mMinute As String
Private mShotFile As String
Private mShot As StdPicture
Private mCursorOff As Boolean
Private mSeenDone As String     ' ;cup id=Timer it was first seen finished;   and the vote: ;vote=...;

Private Sub Form_Load()
    mW = GetSystemMetrics(0)
    mH = GetSystemMetrics(1)
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, mW, mH, SWP_SHOWWINDOW
    ShowCursor 0
    mCursorOff = True
    mSlide = 0
    NextSlide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picSlide.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    mDirty = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrSlide.Enabled = False
    If mCursorOff Then ShowCursor 1
    mCursorOff = False
    Set mShot = Nothing
    mShotFile = ""
End Sub

' Esc closes, the arrows and Space change slides
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape, vbKeyF11
        Unload Me
    Case vbKeyRight, vbKeySpace, vbKeyPageDown
        NextSlide
    Case vbKeyLeft, vbKeyPageUp
        mSlide = mSlide - 2
        If mSlide < 0 Then mSlide = 0
        NextSlide
    End Select
End Sub

Private Sub picSlide_Click()
    NextSlide
End Sub

Private Sub picSlide_DblClick()
    Unload Me
End Sub

' Called by the main window when the schedule, the vote or a tournament changes
Public Sub DataChanged()
    mDirty = True
End Sub

Private Sub tmrSlide_Timer()
    mLeft = mLeft - 1
    If mLeft <= 0 Then
        NextSlide
    ElseIf mDirty Or mKind = "vote" Or Format$(Time, "hh:nn") <> mMinute Then
        DrawSlide
    End If
End Sub

' ================================================================ which slides

' The slides there's something to show on: now, sched, cup:<id> and vote, separated by commas
Private Function SlideList() As String
    Dim want As String, s As String, k As Long
    want = "," & frmMain.ProjSlides() & ","
    If InStr(want, ",now,") > 0 Then
        If frmMain.SchedNowIndex() >= 0 Or frmMain.SchedNextIndex() >= 0 Then s = s & ",now"
    End If
    If InStr(want, ",sched,") > 0 And frmMain.SchedCount() > 0 Then s = s & ",sched"
    If InStr(want, ",cup,") > 0 Then
        For k = 0 To gCupCount - 1
            Select Case gCupState(k)
            Case "signup", "drawn", "running"
                s = s & ",cup:" & gCupId(k)
            Case "done"
                If ShowStill("cup" & gCupId(k), CUP_SHOW) Then s = s & ",cup:" & gCupId(k)
            End Select
        Next
    End If
    If InStr(want, ",vote,") > 0 Then
        Select Case frmMain.VoteText("state")
        Case "open"
            mSeenDone = Replace(mSeenDone, ";vote=" & Seen("vote") & ";", ";")
            s = s & ",vote"
        Case "closed"
            If ShowStill("vote", VOTE_SHOW) Then s = s & ",vote"
        End Select
    End If
    If Len(s) > 0 Then s = Mid$(s, 2)
    SlideList = s
End Function

' Something finished (a vote, a tournament) stays on for a while after it was first seen finished
Private Function ShowStill(ByVal key As String, ByVal secs As Long) As Boolean
    Dim t As String, since As Single
    t = Seen(key)
    If Len(t) = 0 Then
        mSeenDone = mSeenDone & ";" & key & "=" & CStr(Int(Timer)) & ";"
        ShowStill = True
        Exit Function
    End If
    since = Timer - Val(t)
    If since < 0 Then since = since + 86400
    ShowStill = (since < secs)
End Function

Private Function Seen(ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(mSeenDone, ";" & key & "=")
    If p = 0 Then Exit Function
    p = p + Len(key) + 2
    q = InStr(p, mSeenDone, ";")
    If q > p Then Seen = Mid$(mSeenDone, p, q - p)
End Function

Private Sub NextSlide()
    Dim items() As String, s As String
    s = SlideList()
    If Len(s) = 0 Then
        mKind = "welcome"
    Else
        items = Split(s, ",")
        s = items(mSlide Mod (UBound(items) + 1))
        If Left$(s, 4) = "cup:" Then
            mKind = "cup"
            mCupId = Mid$(s, 5)
        Else
            mKind = s
        End If
    End If
    mSlide = mSlide + 1
    mLeft = frmMain.ProjSeconds()
    DrawSlide
End Sub

' ================================================================ drawing

Private Function Px(ByVal n As Single) As Long
    Px = CLng(n * mS)
    If Px < 1 Then Px = 1
End Function

' Text at x, y in Arial, size in pixels (at 800 x 600)
Private Sub Txt(ByVal x As Long, ByVal y As Long, ByVal s As String, ByVal size As Single, ByVal col As Long, Optional ByVal bold As Boolean = True, Optional ByVal maxW As Long = 0)
    SetFont "Arial", size, bold
    If maxW > 0 Then s = FitText(picSlide, s, maxW)
    picSlide.ForeColor = col
    picSlide.CurrentX = x
    picSlide.CurrentY = y
    picSlide.Print s;
End Sub

Private Sub SetFont(ByVal nm As String, ByVal size As Single, ByVal bold As Boolean)
    Dim pts As Single
    pts = Px(size) * Screen.TwipsPerPixelY / 20
    If pts < 6 Then pts = 6
    If picSlide.FontName <> nm Then picSlide.FontName = nm
    If Abs(picSlide.FontSize - pts) > 0.4 Then picSlide.FontSize = pts
    If picSlide.FontBold <> bold Then picSlide.FontBold = bold
End Sub

' Text over several lines, word by word, at most maxLines. y moves down past it.
Private Sub Wrap(ByVal x As Long, y As Long, ByVal s As String, ByVal size As Single, ByVal col As Long, ByVal w As Long, ByVal maxLines As Long)
    Dim words() As String, i As Long, cur As String, try As String, n As Long, lh As Long
    SetFont "Arial", size, False
    lh = picSlide.TextHeight("Ag")
    words = Split(Trim$(s), " ")
    For i = 0 To UBound(words)
        If Len(cur) = 0 Then try = words(i) Else try = cur & " " & words(i)
        If picSlide.TextWidth(try) > w And Len(cur) > 0 Then
            n = n + 1
            If n = maxLines Then
                Txt x, y, FitText(picSlide, cur & " " & words(i) & "...", w), size, col, False
                y = y + lh
                Exit Sub
            End If
            Txt x, y, cur, size, col, False
            y = y + lh
            cur = words(i)
        Else
            cur = try
        End If
    Next
    If Len(cur) > 0 Then
        Txt x, y, cur, size, col, False
        y = y + lh
    End If
End Sub

Private Sub DrawSlide()
    Dim title As String, ic As Integer, foot As String
    mDirty = False
    mMinute = Format$(Time, "hh:nn")
    mW = picSlide.ScaleWidth
    mH = picSlide.ScaleHeight
    mS = mW / 800
    If mH / 600 < mS Then mS = mH / 600
    picSlide.Cls
    Box picSlide, 0, 0, mW, mH, P_BG
    Select Case mKind
    Case "now": DrawNow title, ic, foot
    Case "sched": DrawSched title, ic, foot
    Case "cup": DrawCupSlide title, ic, foot
    Case "vote": DrawVoteSlide title, ic, foot
    Case Else: DrawWelcome title, ic, foot
    End Select
    ' the header: the palm, the slide's icon, its title and the time
    Box picSlide, 0, 0, mW, Px(64), P_BAR
    Box picSlide, 0, Px(64), mW, Px(4), P_ORANGE
    DrawIconSized picSlide, IC_PALM, Px(14), Px(8), Px(48)
    DrawIconSized picSlide, ic, Px(76), Px(16), Px(32)
    SetFont "Arial", 24, True
    Txt mW - Px(14) - picSlide.TextWidth(AmPm(mMinute)), Px(20), AmPm(mMinute), 24, P_PALE
    Txt Px(120), Px(13), title, 34, P_WHITE, True, mW - Px(260)
    If Len(foot) > 0 Then Txt Px(14), mH - Px(40), foot, 22, P_PALE, True, mW - Px(28)
End Sub

Private Sub DrawWelcome(title As String, ic As Integer, foot As String)
    Dim nm As String, sz As Long
    title = "WELCOME"
    ic = IC_PALM
    sz = Px(128)
    DrawIconSized picSlide, IC_PALM, (mW - sz) \ 2, Px(150), sz
    nm = frmMain.PalmsServerName()
    If Len(nm) = 0 Then nm = "Palms Arcade"
    SetFont "Arial", 44, True
    Txt (mW - picSlide.TextWidth(nm)) \ 2, Px(310), nm, 44, P_YELLOW
    SetFont "Arial", 24, False
    Txt (mW - picSlide.TextWidth("Palms Arcade LAN Party")) \ 2, Px(370), "Palms Arcade LAN Party", 24, P_DIM, False
End Sub

Private Sub DrawNow(title As String, ic As Integer, foot As String)
    Dim i As Long, nx As Long, game As String, nm As String, y As Long, x As Long, w As Long, slot As Long, shot As String
    i = frmMain.SchedNowIndex()
    nx = frmMain.SchedNextIndex()
    If i >= 0 Then
        title = "NOW PLAYING"
        If nx >= 0 Then foot = "UP NEXT: " & frmMain.SchedText(nx, "when") & "  " & frmMain.SchedText(nx, "name")
    Else
        title = "COMING UP"
        i = nx
    End If
    ic = IC_JOYSTICK
    If i < 0 Then Exit Sub
    game = frmMain.SchedText(i, "game")
    nm = frmMain.SchedText(i, "name")
    ' the screenshot on the right, if the cartridge has one
    w = mW - Px(28)
    shot = ""
    If Len(game) > 0 Then shot = LibShotPath(game)
    If Len(shot) > 0 Then
        If shot <> mShotFile Then
            On Error Resume Next
            Set mShot = Nothing
            Set mShot = LoadPicture(shot)
            On Error GoTo 0
            mShotFile = shot
        End If
        If Not mShot Is Nothing Then
            x = mW - Px(14) - Px(326)
            Box picSlide, x, Px(82), Px(326), Px(246), P_ORANGE
            picSlide.PaintPicture mShot, x + Px(3), Px(85), Px(320), Px(240)
            w = x - Px(28)
        End If
    End If
    ' the game's icon and name
    y = Px(84)
    slot = -1
    If Len(game) > 0 Then
        If Len(LibIconPath(game)) > 0 Then slot = GameIconSlot(LibIconPath(game))
    End If
    If slot >= 0 Then
        DrawGameIconSized picSlide, slot, Px(14), y, Px(64)
    Else
        DrawIconSized picSlide, frmMain.SchedIconOf(i), Px(14), y, Px(64)
    End If
    If Len(game) > 0 And LibOnServer(game) Then
        Txt Px(90), y + Px(8), LibName(game), 44, P_YELLOW, True, w - Px(76)
        y = y + Px(72)
        If LibName(game) <> nm Then
            Txt Px(14), y, nm, 30, P_WHITE, True, w
            y = y + Px(38)
        End If
    Else
        Txt Px(90), y + Px(8), nm, 44, P_YELLOW, True, w - Px(76)
        y = y + Px(72)
    End If
    If title = "NOW PLAYING" Then
        Txt Px(14), y, "Started " & frmMain.SchedText(i, "when"), 24, P_DIM, False
    Else
        Txt Px(14), y, "Starts " & frmMain.SchedText(i, "when"), 24, P_DIM, False
    End If
    y = y + Px(44)
    If Len(frmMain.SchedText(i, "server")) > 0 Then
        Txt Px(14), y, "SERVER", 16, P_LABEL
        y = y + Px(22)
        SetFont "Courier New", 40, True
        picSlide.ForeColor = P_WHITE
        picSlide.CurrentX = Px(14)
        picSlide.CurrentY = y
        picSlide.Print FitText(picSlide, frmMain.SchedText(i, "server"), w);
        y = y + Px(56)
    End If
    If Len(frmMain.SchedText(i, "connect")) > 0 Then
        Txt Px(14), y, "HOW TO CONNECT", 16, P_LABEL
        y = y + Px(22)
        Wrap Px(14), y, frmMain.SchedText(i, "connect"), 24, P_WHITE, w, 4
    ElseIf Len(frmMain.SchedText(i, "note")) > 0 And Len(frmMain.SchedText(i, "server")) = 0 Then
        Wrap Px(14), y, Replace(frmMain.SchedText(i, "note"), "~", " "), 24, P_WHITE, w, 5
    End If
End Sub

Private Sub DrawSched(title As String, ic As Integer, foot As String)
    Dim n As Long, cur As Long, first As Long, i As Long, k As Long, y As Long, col As Long, t As String, more As Long
    title = "TODAY'S SCHEDULE"
    ic = IC_CALENDAR
    n = frmMain.SchedCount()
    cur = frmMain.SchedNowIndex()
    first = 0
    If cur > 0 Then first = cur - 1
    If first > n - 7 Then first = n - 7
    If first < 0 Then first = 0
    y = Px(88)
    For i = first To first + 6
        If i >= n Then Exit For
        col = P_WHITE
        If i = cur Then
            col = P_YELLOW
        ElseIf cur >= 0 And i < cur Then
            col = P_DIM
        End If
        t = frmMain.SchedText(i, "when")
        SetFont "Arial", 30, True
        Txt Px(184) - picSlide.TextWidth(t), y, t, 30, col
        DrawIconSized picSlide, frmMain.SchedIconOf(i), Px(200), y - Px(2), Px(40)
        t = frmMain.SchedText(i, "name")
        If i = cur Then t = Chr$(187) & " " & t
        Txt Px(256), y, t, 30, col, True, mW - Px(270)
        If Len(frmMain.SchedText(i, "server")) > 0 Then
            Txt Px(256), y + Px(36), frmMain.SchedText(i, "server"), 18, P_DIM, False
            y = y + Px(64)
        Else
            y = y + Px(52)
        End If
        k = i
    Next
    more = n - 1 - k
    If more > 0 Then foot = more & " more later today"
End Sub

Private Sub DrawCupSlide(title As String, ic As Integer, foot As String)
    Dim k As Long, r As Long, r0 As Long, rounds As Long, i As Long, n As Long, players() As String, x As Long, y As Long
    Dim cols As Long, cw As Long, bh As Long, fs As Single, topY As Long, playing As Long, sub1 As String, m As Long, cnt As Long
    title = "TOURNAMENT"
    ic = IC_TROPHY
    k = CupIndex(mCupId)
    If k < 0 Then Exit Sub
    title = "TOURNAMENT: " & UCase$(gCupName(k))
    sub1 = gCupGame(k)
    If Len(gCupTime(k)) > 0 Then
        If Len(sub1) > 0 Then sub1 = sub1 & "  "
        sub1 = sub1 & "starts " & gCupTime(k)
    End If
    If gCupState(k) = "signup" Then
        Txt Px(14), Px(84), "SIGN-UPS OPEN", 44, P_YELLOW
        Txt Px(14), Px(140), sub1, 24, P_DIM, False, mW - Px(28)
        players = Split(gCupPlayers(k), ",")
        n = 0
        For i = 0 To UBound(players)
            If Len(players(i)) > 0 Then n = n + 1
        Next
        If n = 1 Then
            Txt Px(14), Px(186), "1 PLAYER SO FAR", 16, P_LABEL
        Else
            Txt Px(14), Px(186), n & " PLAYERS SO FAR", 16, P_LABEL
        End If
        m = 0
        For i = 0 To UBound(players)
            If Len(players(i)) > 0 And m < 32 Then
                Txt Px(14) + (m Mod 4) * Px(194), Px(214) + (m \ 4) * Px(36), players(i), 26, P_WHITE, True, Px(186)
                m = m + 1
            End If
        Next
        foot = "Sign up in Palms LAN Center: Tournaments > Sign Up."
        Exit Sub
    End If
    rounds = gCupRounds(k)
    If rounds < 1 Then Exit Sub
    topY = Px(84)
    If gCupState(k) = "done" And Len(gCupChamp(k)) > 0 Then
        Txt Px(14), topY, "CHAMPION: " & gCupChamp(k), 44, P_YELLOW, True, mW - Px(28)
        topY = topY + Px(60)
    ElseIf Len(sub1) > 0 Then
        Txt Px(14), topY, sub1, 24, P_DIM, False, mW - Px(28)
        topY = topY + Px(36)
    End If
    ' rounds that fit: from the first one with 8 matches or fewer, unless an earlier one is still being played
    r0 = 0
    Do While r0 < rounds - 1 And CupMatches(k, r0) > 8
        r0 = r0 + 1
    Loop
    playing = -1
    For r = 0 To rounds - 1
        For i = 0 To CupMatches(k, r) - 1
            If CupReady(k, r, i) And playing < 0 Then playing = r
        Next
    Next
    cols = rounds - r0
    If playing >= 0 And playing < r0 Then
        r0 = playing
        cols = 1
    End If
    cnt = CupMatches(k, r0)
    If cnt > 16 Then cnt = 16
    fs = 22
    If cnt > 4 Then fs = 18
    If cnt > 8 Then fs = 15
    cw = (mW - Px(28) - (cols - 1) * Px(12)) \ cols
    bh = Px(fs * 2 + 14)
    For r = r0 To r0 + cols - 1
        x = Px(14) + (r - r0) * (cw + Px(12))
        Txt x, topY, UCase$(CupRoundTitle(k, r)), 16, P_LABEL
        n = CupMatches(k, r)
        If n > 16 Then n = 16
        For i = 0 To n - 1
            ' each round's matches spread over the same height as the first column's
            y = topY + Px(26) + ((i * 2 + 1) * cnt \ n) * (bh + Px(6)) \ 2 - bh \ 2
            If y + bh > mH - Px(10) Then Exit For
            DrawMatch k, r, i, x, y, cw, bh, fs
        Next
    Next
End Sub

Private Sub DrawMatch(ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal fs As Single)
    Dim win As String, sc As String
    picSlide.Line (x, y)-(x + w - 1, y + h - 1), P_BAR, B
    win = CupM(k, r, i, "win")
    sc = CupM(k, r, i, "score")
    DrawPlayer CupM(k, r, i, "a"), win = "a", win = "b", x + Px(6), y + Px(4), w - Px(12), fs, ""
    DrawPlayer CupM(k, r, i, "b"), win = "b", win = "a", x + Px(6), y + Px(6) + Px(fs), w - Px(12), fs, sc
End Sub

Private Sub DrawPlayer(ByVal p As String, ByVal won As Boolean, ByVal lost As Boolean, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal fs As Single, ByVal score As String)
    Dim col As Long
    col = P_WHITE
    If won Then col = P_YELLOW
    If lost Then col = P_DIM
    If p = "-" Then
        p = "bye"
        col = P_DIM
    ElseIf Len(p) = 0 Then
        p = "-"
        col = P_DIM
    End If
    If Len(score) > 0 Then p = p & "   " & score
    Txt x, y, p, fs, col, True, w
End Sub

Private Sub DrawVoteSlide(title As String, ic As Integer, foot As String)
    Dim n As Long, i As Long, total As Long, best As Long, y As Long, bw As Long, col As Long, win As Long, v As Long, secsLeft As Long
    ic = IC_STAR
    If frmMain.VoteText("state") = "open" Then
        title = "NEXT GAME VOTE"
    Else
        title = "VOTE RESULT"
    End If
    Txt Px(14), Px(84), frmMain.VoteText("question"), 44, P_YELLOW, True, mW - Px(28)
    n = Val(frmMain.VoteText("count"))
    total = Val(frmMain.VoteText("total"))
    win = Val(frmMain.VoteText("winner"))
    For i = 0 To n - 1
        If frmMain.VoteVotes(i) > best Then best = frmMain.VoteVotes(i)
    Next
    y = Px(160)
    For i = 0 To n - 1
        v = frmMain.VoteVotes(i)
        col = P_WHITE
        If title = "VOTE RESULT" And win = i + 1 Then col = P_YELLOW
        Txt Px(14), y, frmMain.VoteChoice(i), 30, col, True, Px(300)
        bw = 0
        If total > 0 Then bw = (mW - Px(420)) * v \ total
        If bw < Px(3) Then bw = Px(3)
        If v = best And best > 0 Then
            Box picSlide, Px(330), y + Px(6), bw, Px(24), P_YELLOW
        Else
            Box picSlide, Px(330), y + Px(6), bw, Px(24), P_BAR
        End If
        SetFont "Arial", 30, True
        Txt mW - Px(14) - picSlide.TextWidth(CStr(v)), y, CStr(v), 30, P_WHITE
        y = y + Px(50)
    Next
    If frmMain.VoteText("state") = "open" Then
        secsLeft = Val(frmMain.VoteText("left"))
        foot = "Vote in Palms LAN Center: What's On > Vote. Closes in " & (secsLeft \ 60) & ":" & Format$(secsLeft Mod 60, "00") & "."
    ElseIf win > 0 Then
        foot = frmMain.VoteChoice(win - 1) & " won with " & frmMain.VoteVotes(win - 1) & " of " & total & " votes."
    Else
        foot = "Voting closed."
    End If
End Sub
