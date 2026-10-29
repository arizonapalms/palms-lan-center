VERSION 5.00
Begin VB.Form frmVote
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start Vote"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "frmVote"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6705
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "&Start"
      Height          =   375
      Left            =   5145
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddSugg
      Caption         =   "<< &Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Add the selected game to the choices"
      Top             =   2490
      Width           =   1095
   End
   Begin VB.ListBox lstSugg
      Height          =   1425
      Left            =   5400
      TabIndex        =   10
      Top             =   990
      Width           =   2400
   End
   Begin VB.TextBox txtMins
      Height          =   315
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "5"
      Top             =   2610
      Width           =   495
   End
   Begin VB.TextBox txtOpts
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   990
      Width           =   4815
   End
   Begin VB.TextBox txtQ
      Height          =   315
      Left            =   1320
      MaxLength       =   80
      TabIndex        =   1
      Text            =   "Next game?"
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblSugg
      Caption         =   "S&uggested by users:"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label lblMins2
      Caption         =   "minutes"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2655
      Width           =   1095
   End
   Begin VB.Label lblMins
      Caption         =   "Voting c&loses in"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2655
      Width           =   1335
   End
   Begin VB.Label lblOpts
      Caption         =   "&Choices, one per line (2 to 6):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblQ
      Caption         =   "&Question:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   1095
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Start a next-game vote (host).
Option Explicit

Private mSugg() As String

Private Sub Form_Load()
    Dim i As Long, s As String
    If frmMain.SuggCount() >= 2 Then
        For i = 0 To frmMain.SuggCount() - 1
            If i < 6 Then s = s & frmMain.SuggName(i) & vbCrLf
        Next
        txtOpts.Text = s
    Else
        txtOpts.Text = "Quake 2" & vbCrLf & "Unreal Tournament" & vbCrLf & "Counter-Strike 1.6"
    End If
    RefreshSugg
End Sub

' Called again by the main window whenever the suggestions change
Public Sub RefreshSugg()
    Dim i As Long, n As Long
    lstSugg.Clear
    n = frmMain.SuggCount()
    If n > 0 Then
        ReDim mSugg(0 To n - 1)
    Else
        ReDim mSugg(0 To 0)
    End If
    For i = 0 To n - 1
        mSugg(i) = frmMain.SuggName(i)
        lstSugg.AddItem mSugg(i) & " (" & frmMain.SuggVotes(i) & ")"
    Next
    If n = 0 Then lstSugg.AddItem "No suggestions."
    lstSugg.Enabled = (n > 0)
    cmdAddSugg.Enabled = (n > 0)
End Sub

' Add the selected suggestion to the choices, if it isn't there already
Private Sub cmdAddSugg_Click()
    Dim i As Long, s As String
    i = lstSugg.ListIndex
    If i < 0 Or i > UBound(mSugg) Then Exit Sub
    If Len(mSugg(i)) = 0 Then Exit Sub
    s = txtOpts.Text
    If InStr(1, vbCrLf & s & vbCrLf, vbCrLf & mSugg(i) & vbCrLf, vbTextCompare) > 0 Then Exit Sub
    If Len(s) > 0 And Right$(s, 2) <> vbCrLf Then s = s & vbCrLf
    txtOpts.Text = s & mSugg(i)
End Sub

Private Sub lstSugg_DblClick()
    cmdAddSugg_Click
End Sub

Private Sub cmdOK_Click()
    Dim parts() As String, opts As String, n As Long, i As Long, s As String, q As String, mins As Long
    q = Trim$(Replace(txtQ.Text, "|", "/"))
    If Len(q) = 0 Then q = "Next game?"
    parts = Split(Replace(txtOpts.Text, vbCr, ""), vbLf)
    For i = 0 To UBound(parts)
        s = Trim$(Replace(parts(i), "|", "/"))
        If Len(s) > 0 And n < 6 Then
            opts = opts & "|" & Left$(s, 40)
            n = n + 1
        End If
    Next
    If n < 2 Then
        MsgBox "Enter at least 2 choices, one per line.", vbExclamation, "Start Vote"
        txtOpts.SetFocus
        Exit Sub
    End If
    mins = Val(txtMins.Text)
    If mins < 1 Or mins > 60 Then
        MsgBox "Enter a duration from 1 to 60 minutes.", vbExclamation, "Start Vote"
        txtMins.SetFocus
        Exit Sub
    End If
    frmMain.SendLine "VOTESTART|" & mins & "|" & q & opts
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
