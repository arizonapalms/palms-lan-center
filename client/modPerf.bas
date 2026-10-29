Attribute VB_Name = "modPerf"
' Palms LAN Center - Measures redraw times and processor use for Help > Diagnostics.
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As Currency, lpExitTime As Currency, lpKernelTime As Currency, lpUserTime As Currency) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' What gets measured
Public Const PF_MSG As Integer = 0        ' messages from the server
Public Const PF_TIMER As Integer = 1      ' the once-a-second clock
Public Const PF_LOG As Integer = 2
Public Const PF_WHO As Integer = 3
Public Const PF_NEXT As Integer = 4
Public Const PF_TOOLS As Integer = 5
Public Const PF_SIDE As Integer = 6
Public Const PF_STATUS As Integer = 7
Public Const PF_HEADER As Integer = 8
Public Const PF_WHATS As Integer = 9
Public Const PF_SCHED As Integer = 10
Public Const PF_VOTE As Integer = 11
Public Const PF_WHOPAGE As Integer = 12
Public Const PF_WHOBODY As Integer = 13
Public Const PF_CARD As Integer = 14
Public Const PF_TOAST As Integer = 15
Public Const PF_LIBPAGE As Integer = 16
Public Const PF_LIBLIST As Integer = 17
Public Const PF_LIBCARD As Integer = 18
Public Const PF_LIBTICK As Integer = 19   ' keeping an eye on a game installing
Public Const PF_CUPS As Integer = 20
Public Const PF_COUNT As Integer = 21

Public gPerfCount(0 To 20) As Long        ' how many times
Public gPerfTicks(0 To 20) As Currency    ' how long altogether, in performance-counter ticks
Private mStart(0 To 20) As Currency
Private mFreq As Currency

Public Sub PerfStart(ByVal k As Integer)
    QueryPerformanceCounter mStart(k)
End Sub

Public Sub PerfStop(ByVal k As Integer)
    Dim t As Currency
    QueryPerformanceCounter t
    If t > mStart(k) Then gPerfTicks(k) = gPerfTicks(k) + (t - mStart(k))
    gPerfCount(k) = gPerfCount(k) + 1
End Sub

' Counter ticks to milliseconds
Public Function PerfMs(ByVal ticks As Currency) As Double
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    If mFreq <= 0 Then Exit Function
    PerfMs = ticks / mFreq * 1000#
End Function

Public Function PerfName(ByVal k As Integer) As String
    Select Case k
    Case PF_MSG: PerfName = "Server messages"
    Case PF_TIMER: PerfName = "1-second timer"
    Case PF_LOG: PerfName = "Chat: message log"
    Case PF_WHO: PerfName = "Chat: user list"
    Case PF_NEXT: PerfName = "Chat: Coming Up box"
    Case PF_TOOLS: PerfName = "Toolbar"
    Case PF_SIDE: PerfName = "Sidebar"
    Case PF_STATUS: PerfName = "Status bar"
    Case PF_HEADER: PerfName = "Page title"
    Case PF_WHATS: PerfName = "What's On: whole page"
    Case PF_SCHED: PerfName = "What's On: schedule"
    Case PF_VOTE: PerfName = "What's On: vote"
    Case PF_WHOPAGE: PerfName = "Who's Here: whole page"
    Case PF_WHOBODY: PerfName = "Who's Here: list or map"
    Case PF_CARD: PerfName = "Who's Here: card"
    Case PF_TOAST: PerfName = "Notifications"
    Case PF_LIBPAGE: PerfName = "Game Library: whole page"
    Case PF_LIBLIST: PerfName = "Game Library: list"
    Case PF_LIBCARD: PerfName = "Game Library: card"
    Case PF_LIBTICK: PerfName = "Game Library: installation"
    Case PF_CUPS: PerfName = "Tournaments"
    End Select
End Function

' How much processor time the LAN Center has used so far, in milliseconds.
' -1 on Windows 98, which can't say.
Public Function ProcessCpuMs() As Double
    Dim c As Currency, e As Currency, kt As Currency, ut As Currency
    On Error GoTo Nope
    If GetProcessTimes(GetCurrentProcess(), c, e, kt, ut) = 0 Then GoTo Nope
    ' the times come in 100-nanosecond steps; read as Currency they're already milliseconds
    ProcessCpuMs = CDbl(kt + ut)
    Exit Function
Nope:
    ProcessCpuMs = -1
End Function

' A steady clock in milliseconds
Public Function PerfNowMs() As Double
    Dim t As Currency
    QueryPerformanceCounter t
    PerfNowMs = PerfMs(t)
End Function
