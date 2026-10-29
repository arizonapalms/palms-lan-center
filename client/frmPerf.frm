VERSION 5.00
Begin VB.Form frmPerf
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diagnostics"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7200
   LinkTopic       =   "frmPerf"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTab
      Caption         =   "Performance"
      Height          =   345
      Index           =   2
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optTab
      Caption         =   "This PC"
      Height          =   345
      Index           =   1
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optTab
      Caption         =   "Connections"
      Height          =   345
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Timer tmrPerf
      Interval        =   1000
      Left            =   120
      Top             =   7080
   End
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5985
      TabIndex        =   5
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy
      Caption         =   "&Copy Report"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Copy the full report to the Clipboard, to send to the host"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.PictureBox picPerf
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   120
      ScaleMode       =   3  'Pixel
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   570
      Width           =   6960
   End
End
Attribute VB_Name = "frmPerf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Help > Diagnostics: connections, this PC, and performance.
Option Explicit

Private Const SLOTS As Long = 5

Private mPrevCount(0 To 20) As Long
Private mPrevTicks(0 To 20) As Currency
Private mRingCount(0 To 4, 0 To 20) As Long
Private mRingMs(0 To 4, 0 To 20) As Double
Private mRingCpu(0 To 4) As Double        ' processor time used in that second, -1 if unknown
Private mRingWall(0 To 4) As Double       ' how long that "second" really was
Private mSlot As Long
Private mFilled As Long
Private mPrevCpu As Double
Private mPrevWall As Double
Private mText As String                   ' the Performance tab as text, for Copy Report
Private mPC As String                     ' the This PC tab, worked out once
Private mTab As Integer                   ' 0 = Connections, 1 = This PC, 2 = Performance

Private Sub Form_Load()
    Dim k As Integer
    For k = 0 To PF_COUNT - 1
        mPrevCount(k) = gPerfCount(k)
        mPrevTicks(k) = gPerfTicks(k)
    Next
    mPrevCpu = ProcessCpuMs()
    mPrevWall = PerfNowMs()
    mPC = frmMain.DiagPC()
    mTab = 0
    DrawTab
End Sub

Private Sub optTab_Click(Index As Integer)
    mTab = Index
    DrawTab
End Sub

Private Sub DrawTab()
    picPerf.Cls
    Select Case mTab
    Case 0
        DrawReport frmMain.DiagLive()
    Case 1
        DrawReport mPC
    Case Else
        DrawPerf
    End Select
End Sub

' Lines of "[Heading]" or "Label" & vbTab & "Value", as frmMain.DiagLive and DiagPC make them
Private Sub DrawReport(ByVal s As String)
    Dim lines() As String, i As Long, y As Long, p As Long, w As Long, lbl As String, v As String
    w = picPerf.ScaleWidth
    lines = Split(s, vbCrLf)
    y = 6
    For i = 0 To UBound(lines)
        If Left$(lines(i), 1) = "[" Then
            If y > 6 Then y = y + 8
            TextAt picPerf, 8, y, Mid$(lines(i), 2, Len(lines(i)) - 2), C_TITLE, True
            y = y + 16
            Box picPerf, 8, y - 1, w - 16, 1, vb3DShadow
            y = y + 4
        ElseIf Len(lines(i)) > 0 Then
            p = InStr(lines(i), vbTab)
            If p > 0 Then
                lbl = Left$(lines(i), p - 1)
                v = Mid$(lines(i), p + 1)
            Else
                lbl = lines(i)
                v = ""
            End If
            TextAt picPerf, 16, y, lbl, C_SUB
            TextAt picPerf, 140, y, FitText(picPerf, v, w - 140 - 8), vbBlack
            y = y + 15
        End If
    Next
End Sub

' The same lines as plain text, for Copy Report
Private Function ReportText(ByVal s As String) As String
    Dim lines() As String, i As Long, p As Long, out As String
    lines = Split(s, vbCrLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 1) = "[" Then
            out = out & vbCrLf & Mid$(lines(i), 2, Len(lines(i)) - 2) & vbCrLf
        ElseIf Len(lines(i)) > 0 Then
            p = InStr(lines(i), vbTab)
            If p > 0 Then
                out = out & "  " & Left$(lines(i), p - 1) & ": " & Mid$(lines(i), p + 1) & vbCrLf
            Else
                out = out & "  " & lines(i) & vbCrLf
            End If
        End If
    Next
    ReportText = out
End Function

Private Sub tmrPerf_Timer()
    Dim k As Integer, cpu As Double, wall As Double
    cpu = ProcessCpuMs()
    wall = PerfNowMs()
    For k = 0 To PF_COUNT - 1
        mRingCount(mSlot, k) = gPerfCount(k) - mPrevCount(k)
        mRingMs(mSlot, k) = PerfMs(gPerfTicks(k) - mPrevTicks(k))
        mPrevCount(k) = gPerfCount(k)
        mPrevTicks(k) = gPerfTicks(k)
    Next
    If cpu >= 0 And mPrevCpu >= 0 Then
        mRingCpu(mSlot) = cpu - mPrevCpu
    Else
        mRingCpu(mSlot) = -1
    End If
    mRingWall(mSlot) = wall - mPrevWall
    mPrevCpu = cpu
    mPrevWall = wall
    mSlot = (mSlot + 1) Mod SLOTS
    If mFilled < SLOTS Then mFilled = mFilled + 1
    DrawTab
End Sub

Private Sub DrawPerf()
    Dim k As Integer, i As Long, wall As Double, cpu As Double, cpuKnown As Boolean
    Dim n As Double, ms As Double, drawMs As Double, y As Long, s As String
    mText = "Performance, averaged over " & mFilled & " seconds" & vbCrLf
    If mFilled = 0 Then
        TextAt picPerf, 8, 8, "Measuring...", C_SUB
        Exit Sub
    End If
    cpuKnown = True
    For i = 0 To mFilled - 1
        wall = wall + mRingWall(i)
        If mRingCpu(i) < 0 Then
            cpuKnown = False
        Else
            cpu = cpu + mRingCpu(i)
        End If
    Next
    If wall <= 0 Then Exit Sub
    y = 8
    If cpuKnown Then
        s = "Palms LAN Center is using " & Format$(cpu / wall * 100, "0") & "% of the processor."
    Else
        s = "Processor usage is not available on this version of Windows."
    End If
    TextAt picPerf, 8, y, s, vbBlack, True
    mText = mText & s & vbCrLf
    y = y + 24
    TextAt picPerf, 8, y, "Activity", C_SUB, True
    TextAt picPerf, 196, y, "Per sec", C_SUB, True
    TextAt picPerf, 262, y, "ms each", C_SUB, True
    TextAt picPerf, 330, y, "ms/sec", C_SUB, True
    mText = mText & vbCrLf & "Activity" & vbTab & "Per sec" & vbTab & "ms each" & vbTab & "ms/sec" & vbCrLf
    y = y + 16
    Box picPerf, 8, y, picPerf.ScaleWidth - 16, 1, vb3DShadow
    y = y + 5
    For k = 0 To PF_COUNT - 1
        n = 0
        ms = 0
        For i = 0 To mFilled - 1
            n = n + mRingCount(i, k)
            ms = ms + mRingMs(i, k)
        Next
        n = n / (wall / 1000#)
        ms = ms / (wall / 1000#)
        If k <> PF_MSG And k <> PF_TIMER And k <> PF_LIBTICK Then drawMs = drawMs + ms
        If n > 0 Then
            TextAt picPerf, 8, y, PerfName(k), vbBlack
            TextAt picPerf, 196, y, Format$(n, "0.0"), vbBlack
            TextAt picPerf, 262, y, Format$(ms / n, "0.0"), vbBlack
            TextAt picPerf, 330, y, Format$(ms, "0.0"), vbBlack
            mText = mText & PerfName(k) & vbTab & Format$(n, "0.0") & vbTab & Format$(ms / n, "0.0") & vbTab & Format$(ms, "0.0") & vbCrLf
        Else
            TextAt picPerf, 8, y, PerfName(k), C_GREY
            TextAt picPerf, 196, y, "-", C_GREY
            mText = mText & PerfName(k) & vbTab & "-" & vbCrLf
        End If
        y = y + 16
    Next
    Box picPerf, 8, y + 2, picPerf.ScaleWidth - 16, 1, vb3DShadow
    y = y + 8
    s = "Drawing: " & Format$(drawMs, "0") & " ms per second (" & Format$(drawMs / 10, "0") & "% of the time)."
    TextAt picPerf, 8, y, s, vbBlack, True
    mText = mText & vbCrLf & s & vbCrLf
    y = y + 20
    TextAt picPerf, 8, y, "Times for messages and the timer include any drawing they caused.", C_GREY
End Sub

Private Sub cmdCopy_Click()
    Dim keep As Integer, s As String
    ' the Performance text is made while drawing it
    keep = mTab
    mTab = 2
    DrawTab
    mTab = keep
    DrawTab
    s = "Palms LAN Center " & frmMain.AppVersion & " diagnostics, " & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf
    s = s & ReportText(frmMain.DiagLive()) & ReportText(mPC) & vbCrLf & mText
    Clipboard.Clear
    Clipboard.SetText s
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
