VERSION 5.00
Begin VB.Form frmEvent
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Event"
   ClientHeight    =   3780
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
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.PictureBox picIcon
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4335
      ScaleMode       =   3  'Pixel
      TabIndex        =   10
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
      TabIndex        =   7
      Top             =   1890
      Width           =   4335
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
      TabIndex        =   6
      Top             =   1635
      Width           =   4335
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

Private Sub Form_Load()
    Dim i As Integer, t As Date
    For i = 0 To EVENT_KINDS - 1
        cboIcon.AddItem EventLabel(i)
    Next
    cboIcon.ListIndex = 0
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
    frmMain.SendLine "SCHEDADD|" & hm & "|" & EventKey(cboIcon.ListIndex) & "|" & nm
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
