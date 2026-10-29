Attribute VB_Name = "modWheel"
' Palms LAN Center - Mouse wheel scrolling, which Visual Basic 6 doesn't provide. Off inside the VB editor.
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const WM_MOUSEWHEEL As Long = &H20A

Private mOldProc As Long
Private mHwnd As Long

Public Function InVBEditor() As Boolean
    InVBEditor = (App.LogMode = 0)
End Function

Public Sub WheelHook(ByVal hwnd As Long)
    If InVBEditor() Then Exit Sub
    If mOldProc <> 0 Then Exit Sub
    mHwnd = hwnd
    mOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WheelProc)
End Sub

Public Sub WheelUnhook()
    If mOldProc = 0 Then Exit Sub
    SetWindowLong mHwnd, GWL_WNDPROC, mOldProc
    mOldProc = 0
End Sub

' Must be Public so AddressOf can use it
Public Function WheelProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If uMsg = WM_MOUSEWHEEL Then
        ' the wheel distance is the top half of wParam: +120 per notch away from you
        frmMain.OnWheel wParam \ &H10000
        WheelProc = 0
        Exit Function
    End If
    WheelProc = CallWindowProc(mOldProc, hwnd, uMsg, wParam, lParam)
End Function
