VERSION 5.00
Begin VB.Form frmSeat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Seat"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "frmSeat"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.CommandButton cmdNone
      Caption         =   "&No Seat"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1980
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtSeat
      Height          =   315
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1065
      Width           =   735
   End
   Begin VB.PictureBox picIcon
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleMode       =   3  'Pixel
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblStatus
      Height          =   435
      Left            =   1080
      TabIndex        =   8
      Top             =   1470
      UseMnemonic     =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblSeat
      Caption         =   "S&eat:"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1110
      Width           =   615
   End
   Begin VB.Label lblHint
      Caption         =   "Seat numbers look like A01 or B12. Ask the host if unsure."
      Height          =   450
      Left            =   1080
      TabIndex        =   1
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblPrompt
      Caption         =   "Enter your seat number."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Pick or change your seat.
Option Explicit

Private Sub Form_Load()
    DrawIcon32 picIcon, IC_PEOPLE, 0, 0
    txtSeat.Text = frmMain.MySeat
    If Len(txtSeat.Text) = 0 Then Me.Caption = "Select Seat"
    txtSeat.SelStart = 0
    txtSeat.SelLength = Len(txtSeat.Text)
End Sub

Private Sub cmdOK_Click()
    Dim s As String
    s = UCase$(Trim$(txtSeat.Text))
    If Len(s) = 0 Then
        SendSeat "-"
        Exit Sub
    End If
    If Not ((s Like "[A-Z][0-9][0-9]") Or (s Like "[A-Z][A-Z][0-9][0-9]")) Then
        ShowError "Invalid seat number. Seat numbers look like A01 or B12."
        Exit Sub
    End If
    SendSeat s
End Sub

Private Sub cmdNone_Click()
    txtSeat.Text = ""
    SendSeat "-"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SendSeat(ByVal s As String)
    If Not frmMain.IsOnline Then
        ShowError "Not connected to the server. Try again later."
        Exit Sub
    End If
    If Not frmMain.ServerIsNew Then
        ShowError "Changing seats requires Palms Server version 0.4.1 or later."
        Exit Sub
    End If
    lblStatus.ForeColor = vbButtonText
    lblStatus.Caption = "Saving..."
    frmMain.SendLine "SEAT|" & s
End Sub

' The main window passes on the server's answer when it's a no
Public Sub ShowError(ByVal msg As String)
    lblStatus.ForeColor = C_RED
    lblStatus.Caption = msg
    Beep
    On Error Resume Next
    txtSeat.SetFocus
    txtSeat.SelStart = 0
    txtSeat.SelLength = Len(txtSeat.Text)
End Sub
