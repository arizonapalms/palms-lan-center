VERSION 5.00
Begin VB.Form frmAbout
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Palms LAN Center"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmAbout"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2220
      Width           =   1095
   End
   Begin VB.PictureBox picLogo
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleMode       =   3  'Pixel
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblServer
      Caption         =   "Server"
      ForeColor       =   &H00636163&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblLink
      AutoSize        =   -1  'True
      Caption         =   "palms-arcade.com"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Open the Palms Arcade website"
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblCopy
      Caption         =   "(c) 2026 Palms Arcade"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   900
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblVersion
      Caption         =   "Version"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   630
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblTitle
      Caption         =   "Palms LAN Center"
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
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Help > About.
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1
Private Const HOME_PAGE As String = "http://palms-arcade.com/"

Private Sub Form_Load()
    DrawIcon32 picLogo, IC_PALM, 0, 0
    lblVersion.Caption = "Version " & frmMain.AppVersion
    lblCopy.Caption = Chr$(169) & " 2026 Palms Arcade"
    lblServer.Caption = frmMain.ServerText
End Sub

' Pointing at the link shows the hand, like in a web browser
Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandCursor
End Sub

Private Sub lblLink_Click()
    If ShellExecute(Me.hwnd, "open", HOME_PAGE, vbNullString, vbNullString, SW_SHOWNORMAL) <= 32 Then
        MsgBox "Unable to start the web browser. The website address is " & HOME_PAGE, vbInformation, "Palms Arcade"
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
