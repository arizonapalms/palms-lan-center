VERSION 5.00
Begin VB.Form frmAlert
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Palms LAN Center"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "frmAlert"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen
      Caption         =   "&Open"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2100
      Width           =   1095
   End
   Begin VB.PictureBox picIcon
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleMode       =   3  'Pixel
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lblMsg
      Caption         =   "Message"
      Height          =   1095
      Left            =   960
      TabIndex        =   3
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblTitle
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00733410&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   300
      UseMnemonic     =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - The classic pop-up: a box that waits for OK.
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40

Private mKind As String

' nth = how many of these boxes are already open, so each new one sits a little lower
Public Sub ShowAlert(ByVal kind As String, ByVal ic As Integer, ByVal title As String, ByVal msg As String, ByVal nth As Long)
    mKind = kind
    lblTitle.Caption = title
    lblMsg.Caption = msg
    picIcon.Cls
    DrawIcon32 picIcon, ic, 0, 0
    Me.Move (Screen.Width - Me.Width) / 2 + nth * 300, (Screen.Height - Me.Height) / 2 + nth * 300
    ' on its own, so it never drags the main window in front of a game
    LoneWindow Me.hWnd, False, False
    If gInterrupt Then
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        ForceForeground Me.hWnd
    Else
        ' on top, but without taking the keyboard away from whatever you're doing
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End If
End Sub

Private Sub cmdOpen_Click()
    frmMain.ToastClicked mKind
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
