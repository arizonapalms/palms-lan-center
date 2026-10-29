Attribute VB_Name = "modGfx"
' Palms LAN Center - Drawing helpers for the retro look: icons, bevels, text and colours.
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCAND As Long = &H8800C6
Private Const SRCPAINT As Long = &HEE0086
Private Const IDC_HAND As Long = 32649
Private mHandCursor As Long
Private mCheckBmp As Long
Private mCheckBrush As Long

' ---- icon numbers: positions in the icon sheets in PalmsLAN.res
Public Const IC_PALM As Integer = 0
Public Const IC_CHAT As Integer = 1
Public Const IC_CALENDAR As Integer = 2
Public Const IC_JOYSTICK As Integer = 3
Public Const IC_DISC As Integer = 4
Public Const IC_PEOPLE As Integer = 5
Public Const IC_TROPHY As Integer = 6
Public Const IC_MEGAPHONE As Integer = 7
Public Const IC_USER As Integer = 8
Public Const IC_ADMIN As Integer = 9
Public Const IC_SPEAKER As Integer = 10
Public Const IC_SIG4 As Integer = 11
Public Const IC_SIG3 As Integer = 12
Public Const IC_SIG2 As Integer = 13
Public Const IC_SIG0 As Integer = 14
Public Const IC_OK As Integer = 15
Public Const IC_BAD As Integer = 16
Public Const IC_WARN As Integer = 17
Public Const IC_INFO As Integer = 18
Public Const IC_CLOCK As Integer = 19
Public Const IC_STAR As Integer = 20
Public Const IC_PIZZA As Integer = 21
Public Const IC_MOON As Integer = 22
Public Const IC_DOOR As Integer = 23
Public Const IC_PC As Integer = 24
Public Const IC_PCOFF As Integer = 25
Public Const IC_NET As Integer = 26
Public Const IC_REFRESH As Integer = 27
Public Const IC_MAIL As Integer = 28
Public Const IC_SMILEY As Integer = 29
Public Const IC_FOLDER As Integer = 30
Public Const IC_DOWNLOAD As Integer = 31
Public Const IC_BELL As Integer = 32
Public Const IC_BOX As Integer = 33
Public Const IC_BURGER As Integer = 34
Public Const IC_CROSSHAIR As Integer = 35
Public Const IC_ROCKET As Integer = 36
Public Const IC_FLAG As Integer = 37
Public Const IC_CASTLE As Integer = 38
Public Const IC_PLANET As Integer = 39
Public Const IC_TANK As Integer = 40
Public Const IC_SWORD As Integer = 41
Public Const IC_BOMB As Integer = 42
Public Const IC_PLUS As Integer = 43
Public Const IC_PENCIL As Integer = 44
Public Const IC_DICE As Integer = 45
Public Const IC_FLOPPY As Integer = 46
Public Const IC_COPY As Integer = 47
Public Const IC_LOCK As Integer = 48
Public Const IC_RADIO0 As Integer = 49
Public Const IC_RADIO1 As Integer = 50
Public Const IC_CHECK0 As Integer = 51
Public Const IC_CHECK1 As Integer = 52
Public Const IC_DOWN As Integer = 53
Public Const IC_TCLOSE As Integer = 54
Public Const IC_TMIN As Integer = 55
Public Const IC_TMAX As Integer = 56
Public Const IC_HDD As Integer = 57
Public Const IC_REG As Integer = 58
Public Const IC_SHORTCUT As Integer = 59
Public Const IC_ARROW As Integer = 60
Public Const IC_TABLE As Integer = 61
Public Const IC_PLAN As Integer = 62
Public Const IC_ERASER As Integer = 63
Public Const IC_JUKEBOX As Integer = 64
Public Const IC_PRINTER As Integer = 65
Public Const IC_MUTE As Integer = 66

' ---- colours, written as &HBBGGRR (Visual Basic's order). All are High Color safe.
Public Const C_SIDE As Long = &H6B6131         ' sidebar teal      #31616B
Public Const C_SIDE_ON As Long = &H847942      ' selected button   #427984
Public Const C_SIDE_LT As Long = &HA59A6B      ' sidebar bevel lit #6B9AA5
Public Const C_SIDE_DK As Long = &H423C18      ' sidebar bevel dark #183C42
Public Const C_SIDE_DIM As Long = &HB5AA8C     ' sidebar text when signed out #8CAAB5
Public Const C_TITLE As Long = &H733410        ' page title navy   #103473
Public Const C_SUB As Long = &H424142          ' soft text         #424142
Public Const C_GREY As Long = &H636163         ' times, away       #636163
Public Const C_TIP As Long = &HE7FFFF          ' tooltip yellow    #FFFFE7
Public Const C_SUN1 As Long = &H218EF7         ' sunset orange     #F78E21
Public Const C_SUN2 As Long = &H29DBFF         ' sunset yellow     #FFDB29
Public Const C_BADGE As Long = &H2134E7        ' unread badge red  #E73421
Public Const C_GREEN As Long = &H218218        ' #188221
Public Const C_RED As Long = &H101CAD          ' #AD1C10
Public Const C_BLUE As Long = &H9C5908         ' #08599C
Public Const C_PURPLE As Long = &H8C2463       ' #63248C
Public Const C_BROWN As Long = &H8598C&        ' #8C5908
Public Const C_TEAL As Long = &H7B7D00         ' #007D7B
Public Const C_PINK As Long = &H7B26B5         ' #B5267B
Public Const C_OLIVE As Long = &H18694A        ' #4A6918
Public Const C_ORANGE As Long = &H1065C6       ' #C66510
Public Const C_SEL As Long = &H9C5908          ' selected row      #08599C
Public Const C_TOAST1 As Long = &HE7B28C       ' pop-up sky blue   #8CB2E7
Public Const C_TOAST_EDGE As Long = &HBD927B   ' pop-up outline    #7B92BD
Public Const C_TOAST_IN As Long = &HF7E3D6     ' pop-up inner line #D6E3F7

' How many kinds of event the "Add to the schedule" icon list offers
Public Const EVENT_KINDS As Integer = 15

Private mSpr16 As StdPicture
Private mMask16 As StdPicture
Private mSpr32 As StdPicture
Private mMask32 As StdPicture
Private mLoaded As Boolean

' The icon sheets again, copied once into the screen's own colour format:
' 0 = 16x16 icons, 1 = their mask, 2 = 32x32 icons, 3 = their mask
Private mDC(0 To 3) As Long
Private mFast As Boolean

' Game icons from the pak server (the icons folder): each is a 96 x 32 picture with the
' 32 px icon and its mask, then the 16 px icon and its mask along the top. Each one is
' loaded once, into the screen's own colour format, like the icon sheets.
Private mGameFile() As String
Private mGameDC() As Long             ' 0 when the file couldn't be loaded
Private mGameCount As Long

Public Sub LoadGfx()
    On Error GoTo Failed
    Set mSpr16 = LoadResPicture(101, vbResBitmap)
    Set mMask16 = LoadResPicture(102, vbResBitmap)
    Set mSpr32 = LoadResPicture(103, vbResBitmap)
    Set mMask32 = LoadResPicture(104, vbResBitmap)
    mLoaded = True
    mFast = MakeFast(0, mSpr16)
    If mFast Then mFast = MakeFast(1, mMask16)
    If mFast Then mFast = MakeFast(2, mSpr32)
    If mFast Then mFast = MakeFast(3, mMask32)
    Exit Sub
Failed:
    mLoaded = False
End Sub

' Copies an icon sheet, once, into the same colour format as the screen. Drawing an icon
' from that copy is a straight copy. Drawing from the original made Windows convert the
' whole sheet of icons every time, about 50 ms for each big icon on a Pentium III.
Private Function MakeFast(ByVal i As Integer, pic As StdPicture) As Boolean
    Dim bm As BITMAP, screenDC As Long, srcDC As Long, bmp As Long, old As Long
    On Error GoTo Done
    If GetObjectAPI(pic.Handle, Len(bm), bm) = 0 Then Exit Function
    screenDC = GetDC(0)
    mDC(i) = CreateCompatibleDC(screenDC)
    bmp = CreateCompatibleBitmap(screenDC, bm.bmWidth, bm.bmHeight)
    srcDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    If mDC(i) = 0 Or bmp = 0 Or srcDC = 0 Then GoTo Done
    SelectObject mDC(i), bmp
    old = SelectObject(srcDC, pic.Handle)
    If old = 0 Then GoTo Done
    BitBlt mDC(i), 0, 0, bm.bmWidth, bm.bmHeight, srcDC, 0, 0, SRCCOPY
    SelectObject srcDC, old
    DeleteDC srcDC
    MakeFast = True
    Exit Function
Done:
    If srcDC <> 0 Then DeleteDC srcDC
End Function

' A picture copied into a memory DC in the screen's colour format, for drawing fast.
' Returns the DC, or 0 if it couldn't. Only used when the picture is the size we expect.
Private Function PictureDC(pic As StdPicture, ByVal wantW As Long, ByVal wantH As Long) As Long
    Dim bm As BITMAP, screenDC As Long, srcDC As Long, dc As Long, bmp As Long, old As Long
    On Error GoTo Done
    If GetObjectAPI(pic.Handle, Len(bm), bm) = 0 Then Exit Function
    If bm.bmWidth <> wantW Or bm.bmHeight <> wantH Then Exit Function
    screenDC = GetDC(0)
    dc = CreateCompatibleDC(screenDC)
    bmp = CreateCompatibleBitmap(screenDC, bm.bmWidth, bm.bmHeight)
    srcDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    If dc = 0 Or bmp = 0 Or srcDC = 0 Then GoTo Done
    SelectObject dc, bmp
    old = SelectObject(srcDC, pic.Handle)
    If old = 0 Then GoTo Done
    BitBlt dc, 0, 0, bm.bmWidth, bm.bmHeight, srcDC, 0, 0, SRCCOPY
    SelectObject srcDC, old
    DeleteDC srcDC
    PictureDC = dc
    Exit Function
Done:
    If srcDC <> 0 Then DeleteDC srcDC
    If dc <> 0 Then DeleteDC dc
End Function

' Which loaded game icon a file is, loading it the first time. -1 if it can't be drawn.
Public Function GameIconSlot(ByVal path As String) As Long
    Dim i As Long, pic As StdPicture, n As Long
    GameIconSlot = -1
    If Len(path) = 0 Then Exit Function
    For i = 0 To mGameCount - 1
        If mGameFile(i) = path Then
            If mGameDC(i) <> 0 Then GameIconSlot = i
            Exit Function
        End If
    Next
    If mGameCount = 0 Then
        ReDim mGameFile(0 To 15)
        ReDim mGameDC(0 To 15)
    End If
    If mGameCount > UBound(mGameFile) Then
        ReDim Preserve mGameFile(0 To mGameCount + 16)
        ReDim Preserve mGameDC(0 To mGameCount + 16)
    End If
    n = mGameCount
    mGameFile(n) = path
    mGameDC(n) = 0
    mGameCount = mGameCount + 1
    On Error Resume Next
    Set pic = LoadPicture(path)
    If Err.Number = 0 Then
        If Not pic Is Nothing Then mGameDC(n) = PictureDC(pic, 96, 32)
    End If
    On Error GoTo 0
    If mGameDC(n) <> 0 Then GameIconSlot = n
End Function

' New icons have arrived: try again the ones that weren't there before
Public Sub GameIconsRetry()
    Dim i As Long, n As Long
    For i = 0 To mGameCount - 1
        If mGameDC(i) <> 0 Then
            mGameFile(n) = mGameFile(i)
            mGameDC(n) = mGameDC(i)
            n = n + 1
        End If
    Next
    mGameCount = n
End Sub

Public Sub DrawGameIcon(p As PictureBox, ByVal slot As Long, ByVal big As Boolean, ByVal x As Long, ByVal y As Long)
    If slot < 0 Or slot >= mGameCount Then Exit Sub
    If mGameDC(slot) = 0 Then Exit Sub
    If big Then
        BitBlt p.hDC, x, y, 32, 32, mGameDC(slot), 32, 0, SRCAND
        BitBlt p.hDC, x, y, 32, 32, mGameDC(slot), 0, 0, SRCPAINT
        Touch p, x, y, 32, 32
    Else
        BitBlt p.hDC, x, y, 16, 16, mGameDC(slot), 80, 0, SRCAND
        BitBlt p.hDC, x, y, 16, 16, mGameDC(slot), 64, 0, SRCPAINT
        Touch p, x, y, 16, 16
    End If
End Sub

' A Windows 98 progress bar: sunken, filling up with blue blocks
Public Sub ProgressBar(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal pct As Long)
    Dim fill As Long, bx As Long, bw As Long
    If w < 8 Or h < 5 Then Exit Sub
    Box p, x, y, w, h, vbWindowBackground
    Edge p, x, y, w, h, vb3DShadow, vb3DHighlight
    If pct < 0 Then Exit Sub
    If pct > 100 Then pct = 100
    fill = ((w - 4) * pct) \ 100
    bx = 0
    Do While bx < fill
        bw = 7
        If bx + bw > fill Then bw = fill - bx
        Box p, x + 2 + bx, y + 2, bw, h - 4, C_SEL
        bx = bx + 9
    Loop
End Sub

Public Sub DrawIcon16(p As PictureBox, ByVal idx As Integer, ByVal x As Long, ByVal y As Long)
    Dim sx As Long, sy As Long
    If Not mLoaded Or idx < 0 Then Exit Sub
    sx = (idx Mod 8) * 16
    sy = (idx \ 8) * 16
    If mFast Then
        BitBlt p.hDC, x, y, 16, 16, mDC(1), sx, sy, SRCAND
        BitBlt p.hDC, x, y, 16, 16, mDC(0), sx, sy, SRCPAINT
        Touch p, x, y, 16, 16
    Else
        p.PaintPicture mMask16, x, y, 16, 16, sx, sy, 16, 16, vbSrcAnd
        p.PaintPicture mSpr16, x, y, 16, 16, sx, sy, 16, 16, vbSrcPaint
    End If
End Sub

Public Sub DrawIcon32(p As PictureBox, ByVal idx As Integer, ByVal x As Long, ByVal y As Long)
    Dim sx As Long, sy As Long
    If Not mLoaded Or idx < 0 Then Exit Sub
    sx = (idx Mod 8) * 32
    sy = (idx \ 8) * 32
    If mFast Then
        BitBlt p.hDC, x, y, 32, 32, mDC(3), sx, sy, SRCAND
        BitBlt p.hDC, x, y, 32, 32, mDC(2), sx, sy, SRCPAINT
        Touch p, x, y, 32, 32
    Else
        p.PaintPicture mMask32, x, y, 32, 32, sx, sy, 32, 32, vbSrcAnd
        p.PaintPicture mSpr32, x, y, 32, 32, sx, sy, 32, 32, vbSrcPaint
    End If
End Sub

' After drawing straight onto a picture box's hidden copy (with Windows calls rather than
' Visual Basic's own), tell Windows that part of the screen needs repainting
Private Sub Touch(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim rc As RECT
    rc.Left = x
    rc.Top = y
    rc.Right = x + w
    rc.Bottom = y + h
    InvalidateRect p.hWnd, rc, 0
End Sub

' Turns Windows colour names (like vbButtonFace) into real colours
Public Function SysCol(ByVal c As Long) As Long
    If c < 0 Then
        SysCol = GetSysColor(c And &HFF&)
    Else
        SysCol = c
    End If
End Function

' A filled rectangle
Public Sub Box(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal col As Long)
    If w <= 0 Or h <= 0 Then Exit Sub
    p.Line (x, y)-(x + w - 1, y + h - 1), SysCol(col), BF
End Sub

' A 1-pixel frame: top and left in one colour, bottom and right in another
Public Sub Edge(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal tl As Long, ByVal br As Long)
    Box p, x, y, w, 1, tl
    Box p, x, y, 1, h, tl
    Box p, x, y + h - 1, w, 1, br
    Box p, x + w - 1, y, 1, h, br
End Sub

' The classic 2-pixel sunken border
Public Sub Sunken2(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Edge p, x, y, w, h, vb3DShadow, vb3DHighlight
    Edge p, x + 1, y + 1, w - 2, h - 2, vb3DDKShadow, vb3DLight
End Sub

' The checkerboard Windows uses for a button that's switched on.
' Filled in one go with a patterned brush (drawing it dot by dot was slow on old PCs).
Public Sub Checker(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim bits(0 To 7) As Integer, i As Long, rc As RECT, oldText As Long, oldBack As Long
    If w <= 0 Or h <= 0 Then Exit Sub
    If mCheckBrush = 0 Then
        ' an 8 x 8 black-and-white pattern: every other dot
        For i = 0 To 7
            If i Mod 2 = 0 Then
                bits(i) = &HAA
            Else
                bits(i) = &H55
            End If
        Next
        mCheckBmp = CreateBitmap(8, 8, 1, 1, bits(0))
        If mCheckBmp <> 0 Then mCheckBrush = CreatePatternBrush(mCheckBmp)
    End If
    If mCheckBrush = 0 Then
        Box p, x, y, w, h, vb3DLight
        Exit Sub
    End If
    rc.Left = x
    rc.Top = y
    rc.Right = x + w
    rc.Bottom = y + h
    ' a black-and-white pattern draws in the text and background colours
    oldText = SetTextColor(p.hDC, SysCol(vbButtonFace))
    oldBack = SetBkColor(p.hDC, SysCol(vb3DHighlight))
    FillRect p.hDC, rc, mCheckBrush
    SetTextColor p.hDC, oldText
    SetBkColor p.hDC, oldBack
    Touch p, x, y, w, h
End Sub

' Rounds a colour to the nearest High Color (65,536 colour) value
Public Function Q565(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    r = r \ 8
    g = g \ 4
    b = b \ 8
    Q565 = RGB(r * 8 + r \ 4, g * 4 + g \ 16, b * 8 + b \ 4)
End Function

' Step i of n between two colours
Public Function Mix(ByVal c1 As Long, ByVal c2 As Long, ByVal i As Long, ByVal n As Long) As Long
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    If n < 1 Then n = 1
    r1 = c1 And &HFF&
    g1 = (c1 \ &H100&) And &HFF&
    b1 = (c1 \ &H10000) And &HFF&
    r2 = c2 And &HFF&
    g2 = (c2 \ &H100&) And &HFF&
    b2 = (c2 \ &H10000) And &HFF&
    Mix = Q565(r1 + (r2 - r1) * i \ n, g1 + (g2 - g1) * i \ n, b1 + (b2 - b1) * i \ n)
End Function

' The orange-to-yellow stripe under each page title
Public Sub Sunset(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim i As Long, x0 As Long, x1 As Long
    For i = 0 To 15
        x0 = x + (w * i) \ 16
        x1 = x + (w * (i + 1)) \ 16
        Box p, x0, y, x1 - x0, h, Mix(C_SUN1, C_SUN2, i, 15)
    Next
End Sub

' Writes text at a spot. Text is see-through, so whatever is behind shows.
' The font only changes when it has to: switching bold on and off makes Windows
' build a new font each time, which was slow on old PCs. Always leaves bold off.
Public Sub TextAt(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal s As String, ByVal col As Long, Optional ByVal bold As Boolean = False)
    If p.FontBold <> bold Then p.FontBold = bold
    p.ForeColor = col
    p.CurrentX = x
    p.CurrentY = y
    p.Print s;
    If bold Then p.FontBold = False
End Sub

' Writes text in whatever font the picture box has now: for a run of bold text, set FontBold
' once, write it all with this, then switch it off again
Public Sub TextKeep(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal s As String, ByVal col As Long)
    p.ForeColor = col
    p.CurrentX = x
    p.CurrentY = y
    p.Print s;
End Sub

' Shortens text with "..." so it fits. Finds the right length by halving,
' so a long line takes a handful of measurements instead of one per letter.
Public Function FitText(p As PictureBox, ByVal s As String, ByVal maxW As Long) As String
    Dim lo As Long, hi As Long, m As Long
    If p.TextWidth(s) <= maxW Then
        FitText = s
        Exit Function
    End If
    lo = 0
    hi = Len(s) - 1
    Do While lo < hi
        m = (lo + hi + 1) \ 2
        If p.TextWidth(Left$(s, m) & "...") <= maxW Then
            lo = m
        Else
            hi = m - 1
        End If
    Loop
    FitText = Left$(s, lo) & "..."
End Function

' A group box with an etched line and an icon in its title
Public Sub GroupBox(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal ic As Integer, ByVal title As String)
    Dim tw As Long
    Edge p, x + 1, y + 1, w - 1, h - 1, vb3DHighlight, vb3DHighlight
    Edge p, x, y, w - 1, h - 1, vb3DShadow, vb3DShadow
    tw = p.TextWidth(title)
    If ic >= 0 Then
        Box p, x + 7, y - 8, 3 + 16 + 4 + tw + 3, 16, vbButtonFace
        DrawIcon16 p, ic, x + 10, y - 8
        TextAt p, x + 10 + 16 + 4, y - 6, title, vbButtonText
    Else
        Box p, x + 7, y - 7, tw + 6, 14, vbButtonFace
        TextAt p, x + 10, y - 6, title, vbButtonText
    End If
End Sub

' A yellow tooltip-style box with a thin black border
Public Sub TipBox(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Box p, x, y, w, h, C_TIP
    Edge p, x, y, w, h, vbBlack, vbBlack
End Sub

' The pointing-hand mouse pointer, like over a link. Call it from MouseMove.
Public Sub HandCursor()
    If mHandCursor = 0 Then mHandCursor = LoadCursor(0, IDC_HAND)
    If mHandCursor <> 0 Then SetCursor mHandCursor
End Sub

' True when the mouse pointer is over this window
Public Function MouseIsOver(ByVal hwnd As Long) As Boolean
    MouseIsOver = (WindowUnderMouse() = hwnd)
End Function

' The window the mouse pointer is over right now
Public Function WindowUnderMouse() As Long
    Dim pt As POINTAPI
    GetCursorPos pt
    WindowUnderMouse = WindowFromPoint(pt.x, pt.y)
End Function

' ---- What's On helpers

' Minutes since midnight. Times before 6 AM count as after midnight of the same party.
Public Function PartyMinutes(ByVal hm As String) As Long
    Dim p As Long, m As Long
    p = InStr(hm, ":")
    If p < 2 Then Exit Function
    m = Val(Left$(hm, p - 1)) * 60 + Val(Mid$(hm, p + 1))
    If m < 360 Then m = m + 1440
    PartyMinutes = m
End Function

' "20:30" becomes "8:30 PM"
Public Function AmPm(ByVal hm As String) As String
    Dim p As Long
    p = InStr(hm, ":")
    If p < 2 Then
        AmPm = hm
    ElseIf Val(Left$(hm, p - 1)) < 12 Then
        AmPm = ShortTime(hm) & " AM"
    Else
        AmPm = ShortTime(hm) & " PM"
    End If
End Function

' The icon for a schedule event or a message from the server
Public Function EventIcon(ByVal key As String) As Integer
    Select Case LCase$(key)
    Case "trophy": EventIcon = IC_TROPHY
    Case "crosshair": EventIcon = IC_CROSSHAIR
    Case "rocket": EventIcon = IC_ROCKET
    Case "flag": EventIcon = IC_FLAG
    Case "castle": EventIcon = IC_CASTLE
    Case "planet": EventIcon = IC_PLANET
    Case "tank": EventIcon = IC_TANK
    Case "sword": EventIcon = IC_SWORD
    Case "bomb": EventIcon = IC_BOMB
    Case "pizza": EventIcon = IC_PIZZA
    Case "burger": EventIcon = IC_BURGER
    Case "moon": EventIcon = IC_MOON
    Case "door": EventIcon = IC_DOOR
    Case "megaphone": EventIcon = IC_MEGAPHONE
    Case "joystick": EventIcon = IC_JOYSTICK
    Case "calendar": EventIcon = IC_CALENDAR
    Case "disc": EventIcon = IC_DISC
    Case "plan": EventIcon = IC_PLAN
    Case "jukebox": EventIcon = IC_JUKEBOX
    Case Else: EventIcon = IC_STAR
    End Select
End Function

' The choices in the "Add to the schedule" icon list (0 to EVENT_KINDS - 1)
Public Function EventKey(ByVal i As Integer) As String
    Select Case i
    Case 0: EventKey = "trophy"
    Case 1: EventKey = "crosshair"
    Case 2: EventKey = "rocket"
    Case 3: EventKey = "flag"
    Case 4: EventKey = "castle"
    Case 5: EventKey = "planet"
    Case 6: EventKey = "tank"
    Case 7: EventKey = "sword"
    Case 8: EventKey = "bomb"
    Case 9: EventKey = "pizza"
    Case 10: EventKey = "burger"
    Case 11: EventKey = "moon"
    Case 12: EventKey = "star"
    Case 13: EventKey = "door"
    Case Else: EventKey = "megaphone"
    End Select
End Function

Public Function EventLabel(ByVal i As Integer) As String
    Select Case i
    Case 0: EventLabel = "Tournament"
    Case 1: EventLabel = "Team shooter"
    Case 2: EventLabel = "Arena shooter"
    Case 3: EventLabel = "Capture the flag"
    Case 4: EventLabel = "Strategy"
    Case 5: EventLabel = "Space strategy"
    Case 6: EventLabel = "War game"
    Case 7: EventLabel = "Dungeon crawl"
    Case 8: EventLabel = "Party game"
    Case 9: EventLabel = "Food"
    Case 10: EventLabel = "BBQ or burgers"
    Case 11: EventLabel = "Late night"
    Case 12: EventLabel = "Highlight"
    Case 13: EventLabel = "Doors open or close"
    Case Else: EventLabel = "Announcement"
    End Select
End Function

' Everyone gets their own name colour, the same on every PC
Public Function NickColor(ByVal nick As String) As Long
    Dim i As Long, h As Long
    nick = LCase$(nick)
    For i = 1 To Len(nick)
        h = (h * 31 + Asc(Mid$(nick, i, 1))) Mod 9973
    Next
    Select Case h Mod 9
    Case 0: NickColor = C_RED
    Case 1: NickColor = C_GREEN
    Case 2: NickColor = C_BLUE
    Case 3: NickColor = C_PURPLE
    Case 4: NickColor = C_BROWN
    Case 5: NickColor = C_TEAL
    Case 6: NickColor = C_PINK
    Case 7: NickColor = C_OLIVE
    Case Else: NickColor = C_ORANGE
    End Select
End Function

' "19:05" becomes "7:05", like the clock on the taskbar
Public Function ShortTime(ByVal hm As String) As String
    Dim p As Long, h As Long
    p = InStr(hm, ":")
    If p < 2 Then
        ShortTime = hm
        Exit Function
    End If
    h = Val(Left$(hm, p - 1)) Mod 12
    If h = 0 Then h = 12
    ShortTime = CStr(h) & Mid$(hm, p)
End Function
