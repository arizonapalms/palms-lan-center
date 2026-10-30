VERSION 5.00
Begin VB.Form frmProjSet
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projector Settings"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "frmProjSet"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3060
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Frame fraSlides
      Caption         =   "Slides"
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   4335
      Begin VB.CheckBox chkVote
         Caption         =   "&Vote (while one is open, and its result)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CheckBox chkCup
         Caption         =   "&Tournaments (brackets and results)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   3975
      End
      Begin VB.CheckBox chkSched
         Caption         =   "&Schedule (today's events)"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.CheckBox chkNow
         Caption         =   "&Now Playing (game, server, how to connect)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   3975
      End
   End
   Begin VB.TextBox txtSecs
      Height          =   315
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblWeb
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label lblSecs2
      Caption         =   "seconds"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   285
      Width           =   1335
   End
   Begin VB.Label lblSecs
      Caption         =   "S&how each slide for:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1935
   End
End
Attribute VB_Name = "frmProjSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Host > Projector Settings: seconds per slide and which slides, for every projector.
Option Explicit

Private Sub Form_Load()
    Dim s As String
    txtSecs.Text = CStr(frmMain.ProjSeconds())
    s = "," & frmMain.ProjSlides() & ","
    chkNow.Value = Tick(InStr(s, ",now,") > 0)
    chkSched.Value = Tick(InStr(s, ",sched,") > 0)
    chkCup.Value = Tick(InStr(s, ",cup,") > 0)
    chkVote.Value = Tick(InStr(s, ",vote,") > 0)
    If Len(frmMain.ProjUrl()) > 0 Then
        lblWeb.Caption = "Web page: " & frmMain.ProjUrl() & vbCrLf & "Open it in Internet Explorer and press F11 for full screen."
    End If
End Sub

Private Function Tick(ByVal b As Boolean) As Integer
    If b Then Tick = vbChecked Else Tick = vbUnchecked
End Function

Private Sub cmdOK_Click()
    Dim secs As Long, s As String
    If Not IsNumeric(Trim$(txtSecs.Text)) Then secs = 0 Else secs = Val(txtSecs.Text)
    If secs < 5 Or secs > 300 Then
        MsgBox "Enter a number of seconds from 5 to 300.", vbExclamation, "Projector Settings"
        txtSecs.SetFocus
        Exit Sub
    End If
    If chkNow.Value = vbChecked Then s = s & ",now"
    If chkSched.Value = vbChecked Then s = s & ",sched"
    If chkCup.Value = vbChecked Then s = s & ",cup"
    If chkVote.Value = vbChecked Then s = s & ",vote"
    If Len(s) > 0 Then s = Mid$(s, 2)
    frmMain.SendLine "PROJSET|" & secs & "|" & s
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
