Attribute VB_Name = "modMap"
' Palms LAN Center - The seat map: works out where tables and seats go, and draws them.
Option Explicit

Public Type MapTable
    Name As String
    Prefix As String
    PerSide As Long
    Sides As Long
End Type

Public Type SeatMap
    Id As String
    Name As String
    Stage As String
    Cols As Long
    Count As Long
    Tables() As MapTable
End Type

' Where everything goes, and what each seat shows (the caller fills in the last four)
Public Type MapLayout
    SeatCount As Long
    SeatId() As String
    SeatX() As Long
    SeatY() As Long
    SeatW() As Long
    SeatH() As Long
    SeatName() As String
    SeatState() As Integer
    SeatWarn() As Boolean
    SeatSel() As Boolean
    BanCount As Long
    BanText() As String
    BanX() As Long
    BanY() As Long
    BanW() As Long
    BanH() As Long
    BanStage() As Boolean
    Height As Long
    Small As Boolean
End Type

' What a seat shows
Public Const MS_FREE As Integer = 0
Public Const MS_TAKEN As Integer = 1
Public Const MS_YOU As Integer = 2
Public Const MS_HELD As Integer = 3
Public Const MS_AWAY As Integer = 4
Public Const MS_HOST As Integer = 5

' Colours, &HBBGGRR, all High Color safe
Public Const C_MAPBG As Long = &HE7F7FF       ' cream floor      #FFF7E7
Public Const C_WOOD As Long = &H3169A5        ' table            #A56931
Public Const C_WOOD_LT As Long = &H63A2DE     ' table top edge   #DEA263
Public Const C_WOOD_DK As Long = &H103C6B     ' table under edge #6B3C10
Public Const C_STAGE As Long = &H424142       ' stage            #424142

Private Const TABLE_GAP As Long = 16

' ---- reading and writing a map the way the server sends it

' A map sent as  id|name|stage|columns|tables, starting at field i
' (tables look like  Table A,A,4,2~Table B,B,4,2)
Public Sub MapFromFields(f() As String, ByVal i As Long, m As SeatMap)
    Dim parts() As String, bits() As String, k As Long, n As Long, raw As String
    m.Id = FieldAt(f, i)
    m.Name = FieldAt(f, i + 1)
    m.Stage = FieldAt(f, i + 2)
    m.Cols = Val(FieldAt(f, i + 3))
    If m.Cols < 1 Then m.Cols = 1
    If m.Cols > 4 Then m.Cols = 4
    raw = FieldAt(f, i + 4)
    m.Count = 0
    If Len(raw) = 0 Then
        ReDim m.Tables(0 To 0)
        Exit Sub
    End If
    parts = Split(raw, "~")
    ReDim m.Tables(0 To UBound(parts))
    For k = 0 To UBound(parts)
        bits = Split(parts(k), ",")
        If UBound(bits) >= 3 Then
            m.Tables(n).Name = bits(0)
            m.Tables(n).Prefix = UCase$(Trim$(bits(1)))
            m.Tables(n).PerSide = Val(bits(2))
            If m.Tables(n).PerSide < 1 Then m.Tables(n).PerSide = 1
            If m.Tables(n).PerSide > 12 Then m.Tables(n).PerSide = 12
            If Val(bits(3)) = 2 Then
                m.Tables(n).Sides = 2
            Else
                m.Tables(n).Sides = 1
            End If
            n = n + 1
        End If
    Next
    m.Count = n
End Sub

' The tables for the server, like  Table A,A,4,2~Table B,B,4,2
Public Function MapTablesText(m As SeatMap) As String
    Dim t As Long, s As String
    For t = 0 To m.Count - 1
        If t > 0 Then s = s & "~"
        s = s & MapSafe(m.Tables(t).Name) & "," & m.Tables(t).Prefix & "," & m.Tables(t).PerSide & "," & m.Tables(t).Sides
    Next
    MapTablesText = s
End Function

' Takes out the characters the map format uses to split things up
Public Function MapSafe(ByVal s As String) As String
    MapSafe = Trim$(Replace(Replace(Replace(s, "|", "-"), "~", "-"), ",", " "))
End Function

' "A" and 5 make "A05"
Public Function SeatLabel(ByVal prefix As String, ByVal n As Long) As String
    SeatLabel = prefix & Format$(n, "00")
End Function

Private Function FieldAt(f() As String, ByVal i As Long) As String
    If i >= 0 And i <= UBound(f) Then FieldAt = f(i)
End Function

' ---- laying it out

' Works out where every banner and seat goes, for a picture box availW pixels wide.
' small = True for the editor's preview: little seats with just their number.
Public Sub LayoutMap(m As SeatMap, ByVal availW As Long, ByVal small As Boolean, lay As MapLayout)
    Dim t As Long, k As Long, total As Long, cols As Long, c As Long, cw As Long, colX As Long
    Dim y As Long, rowH As Long, sy As Long, tw As Long, th As Long, gap As Long, gapV As Long, bh As Long
    Dim rowW As Long, x0 As Long, per As Long, margin As Long, maxW As Long, colGap As Long
    lay.Small = small
    total = 0
    For t = 0 To m.Count - 1
        total = total + m.Tables(t).PerSide * m.Tables(t).Sides
    Next
    If total < 1 Then total = 1
    ReDim lay.SeatId(0 To total - 1)
    ReDim lay.SeatX(0 To total - 1)
    ReDim lay.SeatY(0 To total - 1)
    ReDim lay.SeatW(0 To total - 1)
    ReDim lay.SeatH(0 To total - 1)
    ReDim lay.SeatName(0 To total - 1)
    ReDim lay.SeatState(0 To total - 1)
    ReDim lay.SeatWarn(0 To total - 1)
    ReDim lay.SeatSel(0 To total - 1)
    ReDim lay.BanText(0 To m.Count)
    ReDim lay.BanX(0 To m.Count)
    ReDim lay.BanY(0 To m.Count)
    ReDim lay.BanW(0 To m.Count)
    ReDim lay.BanH(0 To m.Count)
    ReDim lay.BanStage(0 To m.Count)
    lay.SeatCount = 0
    lay.BanCount = 0
    If small Then
        margin = 4
        th = 16
        gap = 3
        gapV = 2
        bh = 14
        maxW = 36
        colGap = 8
    Else
        margin = 8
        th = 40
        gap = 8
        gapV = 4
        bh = 18
        maxW = 130
        colGap = TABLE_GAP
    End If
    y = margin
    If Len(m.Stage) > 0 Then
        If small Then
            AddBanner lay, m.Stage, margin, y, availW - 2 * margin, 14, True
            y = y + 14 + 8
        Else
            AddBanner lay, m.Stage, margin, y, availW - 2 * margin, 26, True
            y = y + 26 + 14
        End If
    End If
    cols = m.Cols
    If cols > m.Count Then cols = m.Count
    If cols < 1 Then cols = 1
    cw = (availW - 2 * margin - (cols - 1) * colGap) \ cols
    rowH = 0
    For t = 0 To m.Count - 1
        c = t Mod cols
        If c = 0 And t > 0 Then
            y = y + rowH + th \ 2 + 6
            rowH = 0
        End If
        colX = margin + c * (cw + colGap)
        per = m.Tables(t).PerSide
        If per < 1 Then per = 1
        tw = (cw - (per - 1) * gap) \ per
        If tw > maxW Then tw = maxW
        If tw < 10 Then tw = 10
        rowW = per * tw + (per - 1) * gap
        x0 = colX + (cw - rowW) \ 2
        If x0 < colX Then x0 = colX
        sy = y
        For k = 1 To per
            AddSeat lay, SeatLabel(m.Tables(t).Prefix, k), x0 + (k - 1) * (tw + gap), sy, tw, th
        Next
        sy = sy + th + gapV
        AddBanner lay, UCase$(m.Tables(t).Name), x0, sy, rowW, bh, False
        sy = sy + bh
        If m.Tables(t).Sides = 2 Then
            sy = sy + gapV
            For k = 1 To per
                AddSeat lay, SeatLabel(m.Tables(t).Prefix, per + k), x0 + (k - 1) * (tw + gap), sy, tw, th
            Next
            sy = sy + th
        End If
        If sy - y > rowH Then rowH = sy - y
    Next
    lay.Height = y + rowH + margin
End Sub

Private Sub AddSeat(lay As MapLayout, ByVal id As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim i As Long
    i = lay.SeatCount
    If i > UBound(lay.SeatId) Then Exit Sub
    lay.SeatId(i) = id
    lay.SeatX(i) = x
    lay.SeatY(i) = y
    lay.SeatW(i) = w
    lay.SeatH(i) = h
    lay.SeatName(i) = ""
    lay.SeatState(i) = MS_FREE
    lay.SeatWarn(i) = False
    lay.SeatSel(i) = False
    lay.SeatCount = i + 1
End Sub

Private Sub AddBanner(lay As MapLayout, ByVal txt As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal isStage As Boolean)
    Dim i As Long
    i = lay.BanCount
    If i > UBound(lay.BanText) Then Exit Sub
    lay.BanText(i) = txt
    lay.BanX(i) = x
    lay.BanY(i) = y
    lay.BanW(i) = w
    lay.BanH(i) = h
    lay.BanStage(i) = isStage
    lay.BanCount = i + 1
End Sub

' The seat under a point (in map pixels, before scrolling), or -1
Public Function SeatAt(lay As MapLayout, ByVal px As Single, ByVal py As Single) As Long
    Dim i As Long
    SeatAt = -1
    For i = 0 To lay.SeatCount - 1
        If px >= lay.SeatX(i) And px < lay.SeatX(i) + lay.SeatW(i) And py >= lay.SeatY(i) And py < lay.SeatY(i) + lay.SeatH(i) Then
            SeatAt = i
            Exit Function
        End If
    Next
End Function

' The seat called id (like A05), or -1
Public Function SeatIndex(lay As MapLayout, ByVal id As String) As Long
    Dim i As Long
    SeatIndex = -1
    If Len(id) = 0 Then Exit Function
    For i = 0 To lay.SeatCount - 1
        If lay.SeatId(i) = id Then
            SeatIndex = i
            Exit Function
        End If
    Next
End Function

' ---- drawing it

' Draws the map onto p, scrolled up by yOff pixels
Public Sub DrawMap(p As PictureBox, lay As MapLayout, ByVal yOff As Long)
    Dim i As Long
    For i = 0 To lay.BanCount - 1
        If lay.BanY(i) - yOff < p.ScaleHeight And lay.BanY(i) + lay.BanH(i) - yOff > 0 Then DrawBanner p, lay, i, yOff
    Next
    For i = 0 To lay.SeatCount - 1
        If lay.SeatY(i) - yOff < p.ScaleHeight And lay.SeatY(i) + lay.SeatH(i) - yOff > 0 Then DrawSeat p, lay, i, yOff
    Next
End Sub

Private Sub DrawBanner(p As PictureBox, lay As MapLayout, ByVal i As Long, ByVal yOff As Long)
    Dim x As Long, y As Long, w As Long, h As Long, s As String, tw As Long, bold As Boolean
    x = lay.BanX(i)
    y = lay.BanY(i) - yOff
    w = lay.BanW(i)
    h = lay.BanH(i)
    bold = Not lay.Small
    p.FontBold = bold
    If lay.BanStage(i) Then
        Box p, x, y, w, h, C_STAGE
        s = FitText(p, lay.BanText(i), w - 40)
        tw = p.TextWidth(s)
        If lay.Small Then
            TextAt p, x + (w - tw) \ 2, y + 1, s, vbWhite, bold
        Else
            DrawIcon16 p, IC_STAR, x + (w - tw) \ 2 - 22, y + (h - 16) \ 2
            TextAt p, x + (w - tw) \ 2, y + (h - 13) \ 2, s, vbWhite, bold
        End If
    Else
        Box p, x, y, w, h, C_WOOD
        Box p, x, y, w, 1, C_WOOD_LT
        Box p, x, y + h - 1, w, 1, C_WOOD_DK
        s = FitText(p, lay.BanText(i), w - 8)
        tw = p.TextWidth(s)
        TextAt p, x + (w - tw) \ 2, y + (h - 13) \ 2, s, vbWhite, bold
    End If
    p.FontBold = False
End Sub

Private Sub DrawSeat(p As PictureBox, lay As MapLayout, ByVal i As Long, ByVal yOff As Long)
    Dim x As Long, y As Long, w As Long, h As Long, st As Integer, sel As Boolean
    Dim ic As Integer, l1 As String, col As Long, col2 As Long, tx As Long, d As Long
    x = lay.SeatX(i)
    y = lay.SeatY(i) - yOff
    w = lay.SeatW(i)
    h = lay.SeatH(i)
    st = lay.SeatState(i)
    sel = lay.SeatSel(i)
    If lay.Small Then
        ' the editor's preview: a plain box with the seat number
        Box p, x, y, w, h, vbWhite
        Edge p, x, y, w, h, vb3DShadow, vb3DShadow
        TextAt p, x + (w - p.TextWidth(lay.SeatId(i))) \ 2, y + 1, FitText(p, lay.SeatId(i), w - 2), C_SUB
        Exit Sub
    End If
    ' the tile: flat and cream when nobody's there, a raised button when someone is
    If sel Then
        Box p, x, y, w, h, C_SEL
        Edge p, x, y, w, h, vb3DDKShadow, vb3DHighlight
        d = 1
    ElseIf st = MS_FREE Or st = MS_HELD Then
        Box p, x, y, w, h, C_MAPBG
        Edge p, x, y, w, h, vb3DShadow, vb3DShadow
        Edge p, x + 1, y + 1, w - 2, h - 2, vbWhite, vbWhite
    Else
        Box p, x, y, w, h, vbButtonFace
        Edge p, x, y, w, h, vb3DHighlight, vb3DDKShadow
        Edge p, x + 1, y + 1, w - 2, h - 2, vb3DLight, vb3DShadow
    End If
    Select Case st
    Case MS_FREE
        ic = IC_PCOFF
        l1 = "Free"
    Case MS_HELD
        ic = IC_LOCK
        l1 = "Reserved"
    Case MS_YOU
        ic = IC_USER
        l1 = "You"
    Case MS_AWAY
        ic = IC_MOON
        l1 = lay.SeatName(i)
    Case MS_HOST
        ic = IC_ADMIN
        l1 = lay.SeatName(i)
    Case Else
        ic = IC_PC
        l1 = lay.SeatName(i)
    End Select
    If sel Then
        col = vbWhite
        col2 = vbWhite
    ElseIf st = MS_FREE Or st = MS_HELD Or st = MS_AWAY Then
        col = C_GREY
        col2 = C_GREY
    Else
        col = vbBlack
        col2 = C_SUB
    End If
    If w >= 64 Then
        DrawIcon16 p, ic, x + 6 + d, y + (h - 16) \ 2 + d
        tx = x + 26 + d
    Else
        tx = x + 4 + d
    End If
    p.FontBold = True
    l1 = FitText(p, l1, x + w - tx - 4)
    p.FontBold = False
    TextAt p, tx, y + 5 + d, l1, col, True
    TextAt p, tx, y + 21 + d, lay.SeatId(i), col2
    If lay.SeatWarn(i) Then DrawIcon16 p, IC_WARN, x + w - 19, y + 3
End Sub

' The key along the bottom: what each kind of seat looks like
Public Sub DrawMapKey(p As PictureBox, ByVal x As Long, ByVal y As Long)
    x = KeyItem(p, x, y, IC_USER, "You")
    x = KeyItem(p, x, y, IC_PC, "Taken")
    x = KeyItem(p, x, y, IC_PCOFF, "Free")
    x = KeyItem(p, x, y, IC_LOCK, "Reserved")
    x = KeyItem(p, x, y, IC_MOON, "Away")
    x = KeyItem(p, x, y, IC_WARN, "Network issue")
End Sub

Private Function KeyItem(p As PictureBox, ByVal x As Long, ByVal y As Long, ByVal ic As Integer, ByVal s As String) As Long
    DrawIcon16 p, ic, x, y
    TextAt p, x + 19, y + 2, s, vbBlack
    KeyItem = x + 19 + p.TextWidth(s) + 14
End Function
