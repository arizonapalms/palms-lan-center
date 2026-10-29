VERSION 5.00
Begin VB.Form frmToast
   BorderStyle     =   0  'None
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "frmToast"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleMode       =   3  'Pixel
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrToast
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picBody
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   132
      Left            =   0
      ScaleMode       =   3  'Pixel
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   200
   End
End
Attribute VB_Name = "frmToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - The Messenger-style pop-up in the bottom-right corner.
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SPI_GETWORKAREA As Long = 48
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const TW As Long = 200          ' size in pixels
Private Const TH As Long = 132
Private Const STEP_PX As Long = 12      ' how far it slides each tick
Private Const SLIDE_MS As Long = 25     ' ticks while sliding
Private Const WAIT_MS As Long = 200     ' ticks while it just sits there (saves the processor)
Private Const XB_L As Long = 180        ' the close button: left, top and size
Private Const XB_T As Long = 4
Private Const XB_S As Long = 16

Private mKind As String
Private mIcon As Integer
Private mTitle As String
Private mMsg As String
Private mSlot As Long
Private mX As Long
Private mBase As Long       ' where the bottom edge sits on screen when fully up
Private mVis As Long        ' how many pixels are showing
Private mPhase As Integer   ' 1 = sliding up, 2 = waiting, 3 = sliding down
Private mUntil As Long
Private mLife As Long
Private mOverText As Boolean    ' pointing at the pop-up (not the x)
Private mOverX As Boolean       ' pointing at the x
Private mDownX As Boolean       ' the x is being pressed

Public Sub Popup(ByVal kind As String, ByVal ic As Integer, ByVal title As String, ByVal msg As String, ByVal secs As Long)
    Dim rc As RECT, i As Long
    mKind = kind
    mIcon = ic
    mTitle = title
    mMsg = msg
    mLife = secs * 1000
    ' take the lowest free spot in the stack
    mSlot = 3
    For i = 3 To 0 Step -1
        If Not gToastSlot(i) Then mSlot = i
    Next
    gToastSlot(mSlot) = True
    picBody.Move 0, 0, TW, TH
    DrawToast
    SystemParametersInfo SPI_GETWORKAREA, 0, rc, 0
    mX = rc.Right - TW - 8
    mBase = rc.Bottom - 8 - mSlot * (TH + 8)
    mVis = 0
    ' on its own, so it never drags the main window in front of a game
    LoneWindow Me.hWnd, True, Not gInterrupt
    Place
    If gInterrupt Then ForceForeground Me.hWnd
    SetPhase 1
    tmrToast.Enabled = True
End Sub

' Shows the top mVis pixels, sitting on the base line, so it looks like it rises from behind it
Private Sub Place()
    Dim h As Long
    h = mVis
    If h < 0 Then h = 0
    If h > TH Then h = TH
    SetWindowRgn Me.hWnd, CreateRectRgn(0, 0, TW, h), 1
    SetWindowPos Me.hWnd, HWND_TOPMOST, mX, mBase - h, TW, TH, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub SetPhase(ByVal ph As Integer)
    mPhase = ph
    If ph = 2 Then
        tmrToast.Interval = WAIT_MS
    Else
        tmrToast.Interval = SLIDE_MS
    End If
End Sub

Private Sub CloseNow()
    tmrToast.Enabled = False
    Unload Me
End Sub

Private Sub tmrToast_Timer()
    ' the mouse left without telling us: un-highlight
    If mOverText Or mOverX Then
        If Not mDownX And WindowUnderMouse() <> picBody.hWnd Then
            mOverText = False
            mOverX = False
            DrawToast
        End If
    End If
    Select Case mPhase
    Case 1
        mVis = mVis + STEP_PX
        If mVis >= TH Then
            mVis = TH
            SetPhase 2
            mUntil = GetTickCount() + mLife
        End If
        Place
    Case 2
        If mOverText Or mOverX Then
            ' pointing at it keeps it up
            If mUntil - GetTickCount() < 2500 Then mUntil = GetTickCount() + 2500
        ElseIf GetTickCount() - mUntil >= 0 Then
            SetPhase 3
        End If
    Case 3
        mVis = mVis - STEP_PX
        If mVis <= 0 Then
            CloseNow
            Exit Sub
        End If
        Place
    End Select
End Sub

Private Function InX(ByVal px As Single, ByVal py As Single) As Boolean
    InX = (px >= XB_L And px < XB_L + XB_S And py >= XB_T And py < XB_T + XB_S)
End Function

Private Sub picBody_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim overX As Boolean, overText As Boolean
    overX = InX(X, Y)
    overText = Not overX And X >= 0 And Y >= 0 And X < TW And Y < TH
    If mDownX Then overText = False
    If overText Then HandCursor
    If overX <> mOverX Or overText <> mOverText Then
        mOverX = overX
        mOverText = overText
        DrawToast
    End If
    If mPhase = 3 Then SetPhase 1        ' it was leaving: come back
End Sub

Private Sub picBody_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If InX(X, Y) Then
        mDownX = True
        DrawToast
    Else
        HandCursor
    End If
End Sub

Private Sub picBody_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If mDownX Then
        mDownX = False
        If InX(X, Y) Then
            CloseNow
        Else
            DrawToast
        End If
        Exit Sub
    End If
    If X >= 0 And Y >= 0 And X < TW And Y < TH Then
        frmMain.ToastClicked mKind
        CloseNow
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrToast.Enabled = False
    gToastSlot(mSlot) = False
End Sub

Private Sub DrawToast()
    Dim i As Long, k As Long, y0 As Long, y1 As Long, x0 As Long, x1 As Long
    Dim rows(0 To 3) As String, n As Long, y As Long, tw2 As Long, col As Long, s As String, d As Long
    PerfStart PF_TOAST
    picBody.Cls
    ' sky blue at the top fading to white
    For i = 0 To 15
        y0 = (TH * i) \ 16
        y1 = (TH * (i + 1)) \ 16
        k = i
        If k > 11 Then k = 11
        Box picBody, 0, y0, TW, y1 - y0, Mix(C_TOAST1, vbWhite, k, 11)
    Next
    ' the frame
    Edge picBody, 0, 0, TW, TH, C_TOAST_EDGE, C_TOAST_EDGE
    Edge picBody, 1, 1, TW - 2, TH - 2, vbWhite, vbWhite
    Edge picBody, 2, 2, TW - 4, TH - 4, vbWhite, vbWhite
    Edge picBody, 3, 3, TW - 6, TH - 6, C_TOAST_IN, C_TOAST_IN
    ' the coloured stripe along the bottom: Palms sunset, orange, teal, green
    y = TH - 4 - 3
    x0 = 4
    x1 = 4 + ((TW - 8) * 55) \ 100
    Box picBody, x0, y, x1 - x0, 3, C_SUN2
    x0 = x1
    x1 = 4 + ((TW - 8) * 70) \ 100
    Box picBody, x0, y, x1 - x0, 3, C_SUN1
    x0 = x1
    x1 = 4 + ((TW - 8) * 85) \ 100
    Box picBody, x0, y, x1 - x0, 3, C_SIDE
    x0 = x1
    x1 = TW - 4
    Box picBody, x0, y, x1 - x0, 3, C_GREEN
    ' a small Palms mark above the stripe
    tw2 = picBody.TextWidth("Palms")
    TextAt picBody, TW - 8 - tw2, TH - 22, "Palms", C_GREY
    DrawIcon16 picBody, IC_PALM, TW - 8 - tw2 - 18, TH - 24
    ' the close button: flat, a raised square when pointed at, pushed in while pressed
    d = 0
    If mDownX And mOverX Then
        Box picBody, XB_L, XB_T, XB_S, XB_S, C_TOAST_IN
        Edge picBody, XB_L, XB_T, XB_S, XB_S, C_TOAST_EDGE, vbWhite
        d = 1
    ElseIf mOverX Or mDownX Then
        Edge picBody, XB_L, XB_T, XB_S, XB_S, vbWhite, C_TOAST_EDGE
    End If
    DrawIcon16 picBody, IC_TCLOSE, XB_L + d, XB_T + d
    ' the big icon
    DrawIcon32 picBody, mIcon, 10, 10
    ' the words, centred like Messenger did it. Blue and underlined while pointed at, like a link.
    If mOverText Then
        col = vbBlue
    Else
        col = vbBlack
    End If
    picBody.FontUnderline = mOverText
    s = FitText(picBody, mTitle, TW - 20)
    TextAt picBody, (TW - picBody.TextWidth(s)) \ 2, 48, s, col
    n = WrapInto(mMsg, TW - 24, rows)
    For i = 0 To n - 1
        TextAt picBody, (TW - picBody.TextWidth(rows(i))) \ 2, 64 + i * 14, rows(i), col
    Next
    picBody.FontUnderline = False
    PerfStop PF_TOAST
End Sub

' Breaks text into up to 3 centred lines. The last line gets "..." if it doesn't all fit.
Private Function WrapInto(ByVal s As String, ByVal maxW As Long, rows() As String) As Long
    Dim words() As String, i As Long, n As Long, cur As String, trial As String
    words = Split(Trim$(s), " ")
    For i = 0 To UBound(words)
        If Len(cur) = 0 Then
            trial = words(i)
        Else
            trial = cur & " " & words(i)
        End If
        If Len(cur) = 0 Or picBody.TextWidth(trial) <= maxW Then
            cur = trial
        Else
            If n = 2 Then
                rows(2) = FitText(picBody, cur & " " & words(i) & "...", maxW)
                WrapInto = 3
                Exit Function
            End If
            rows(n) = cur
            n = n + 1
            cur = words(i)
        End If
    Next
    If Len(cur) > 0 Then
        rows(n) = FitText(picBody, cur, maxW)
        n = n + 1
    End If
    WrapInto = n
End Function
