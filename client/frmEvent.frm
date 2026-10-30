VERSION 5.00
Begin VB.Form frmEvent
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Event"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "frmEvent"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   4590
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   4590
      Width           =   1095
   End
   Begin VB.PictureBox picIcon
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4335
      ScaleMode       =   3  'Pixel
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1178
      Width           =   240
   End
   Begin VB.TextBox txtNote
      Height          =   1095
      Left            =   240
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox txtConnect
      Height          =   315
      Left            =   2280
      MaxLength       =   120
      TabIndex        =   11
      Top             =   2490
      Width           =   2295
   End
   Begin VB.TextBox txtServer
      Height          =   315
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cboGame
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1590
      Width           =   2295
   End
   Begin VB.ComboBox cboIcon
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1140
      Width           =   1935
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   3
      Top             =   690
      Width           =   2295
   End
   Begin VB.TextBox txtTime
      Height          =   315
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblNote
      Caption         =   "&Details (optional):"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2985
      Width           =   4335
   End
   Begin VB.Label lblConnect
      Caption         =   "How to co&nnect:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2535
      Width           =   1935
   End
   Begin VB.Label lblServer
      Caption         =   "Game &server:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2085
      Width           =   1935
   End
   Begin VB.Label lblGame
      Caption         =   "&Game:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1635
      Width           =   1935
   End
   Begin VB.Label lblIcon
      Caption         =   "&Icon:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1185
      Width           =   1935
   End
   Begin VB.Label lblName
      Caption         =   "&Event:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   735
      Width           =   1935
   End
   Begin VB.Label lblTime
      Caption         =   "&Time (24-hour):"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1935
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Add an event to today's schedule (host).
Option Explicit

Private mGameKey() As String     ' the Game list's games: their keys, like quake3/quake3

Private Sub Form_Load()
    Dim i As Integer, t As Date
    For i = 0 To EVENT_KINDS - 1
        cboIcon.AddItem EventLabel(i)
    Next
    cboIcon.ListIndex = 0
    FillGames
    ' older Palms Servers have no game for an event
    If Not frmMain.ServerHasProjector() Then
        lblGame.Visible = False
        cboGame.Visible = False
        lblServer.Visible = False
        txtServer.Visible = False
        lblConnect.Visible = False
        txtConnect.Visible = False
        lblNote.Top = lblNote.Top - 1350
        txtNote.Top = txtNote.Top - 1350
        cmdOK.Top = cmdOK.Top - 1350
        cmdCancel.Top = cmdCancel.Top - 1350
        Me.Height = Me.Height - 1350
    End If
    ' suggest the next half hour
    t = DateAdd("n", 30 - (Minute(Now) Mod 30), Now)
    txtTime.Text = Format$(t, "hh:nn")
    ' older Palms Servers have no event details
    If Not frmMain.ServerHasWhatsOn() Then
        lblNote.Visible = False
        txtNote.Visible = False
        cmdOK.Top = lblNote.Top + 105
        cmdCancel.Top = cmdOK.Top
        Me.Height = Me.Height - 1500
    End If
End Sub

Private Sub FillGames()
    Dim lines() As String, i As Long, p As Long
    cboGame.AddItem "(none)"
    ReDim mGameKey(0 To 0)
    lines = Split(LibGameList(), vbLf)
    For i = 0 To UBound(lines)
        p = InStr(lines(i), vbTab)
        If p > 1 Then
            cboGame.AddItem Mid$(lines(i), p + 1)
            ReDim Preserve mGameKey(0 To cboGame.ListCount - 1)
            mGameKey(cboGame.ListCount - 1) = Left$(lines(i), p - 1)
        End If
    Next
    cboGame.ListIndex = 0
End Sub

' Picking a game names the event after it, if it has no name yet
Private Sub cboGame_Click()
    If cboGame.ListIndex > 0 And Len(Trim$(txtName.Text)) = 0 Then txtName.Text = cboGame.Text
End Sub

Private Sub cboIcon_Click()
    picIcon.Cls
    DrawIcon16 picIcon, EventIcon(EventKey(cboIcon.ListIndex)), 0, 0
End Sub

' Accepts 19:30, 7:30 or 19.30 and gives back "19:30"
Private Function ValidTime(ByVal s As String, hm As String) As Boolean
    Dim p As Long, h As Long, m As Long
    s = Trim$(Replace(s, ".", ":"))
    p = InStr(s, ":")
    If p < 2 Or p > 3 Then Exit Function
    If Len(s) <> p + 2 Then Exit Function
    If Not IsNumeric(Left$(s, p - 1)) Then Exit Function
    If Not IsNumeric(Mid$(s, p + 1)) Then Exit Function
    h = Val(Left$(s, p - 1))
    m = Val(Mid$(s, p + 1))
    If h < 0 Or h > 23 Or m < 0 Or m > 59 Then Exit Function
    hm = Format$(h, "00") & ":" & Format$(m, "00")
    ValidTime = True
End Function

Private Sub cmdOK_Click()
    Dim hm As String, nm As String
    If Not ValidTime(txtTime.Text, hm) Then
        MsgBox "Enter the time in 24-hour format (for example, 19:30 for 7:30 PM).", vbExclamation, "Add Event"
        txtTime.SetFocus
        Exit Sub
    End If
    nm = Trim$(txtName.Text)
    If Len(nm) = 0 Then
        MsgBox "Enter a description for the event (for example, Quake 2 tournament).", vbExclamation, "Add Event"
        txtName.SetFocus
        Exit Sub
    End If
    If txtNote.Visible Then frmMain.SetPendingNote NoteFromText(txtNote.Text)
    If cboGame.Visible And (cboGame.ListIndex > 0 Or Len(Trim$(txtServer.Text)) > 0 Or Len(Trim$(txtConnect.Text)) > 0) Then
        frmMain.SetPendingGame mGameKey(cboGame.ListIndex) & "|" & NoBar(txtServer.Text) & "|" & NoBar(txtConnect.Text)
    End If
    frmMain.SendLine "SCHEDADD|" & hm & "|" & EventKey(cboIcon.ListIndex) & "|" & nm
    Unload Me
End Sub

Private Function NoBar(ByVal s As String) As String
    NoBar = Trim$(Replace(Replace(Replace(s, "|", "/"), vbCr, " "), vbLf, " "))
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
