VERSION 5.00
Begin VB.Form frmEventInfo
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Details"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmEventInfo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConnect
      Height          =   315
      Left            =   1800
      MaxLength       =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   3360
   End
   Begin VB.TextBox txtServer
      Height          =   315
      Left            =   1800
      MaxLength       =   40
      TabIndex        =   3
      Top             =   1290
      Width           =   3360
   End
   Begin VB.ComboBox cboGame
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   3360
   End
   Begin VB.Label lblConnect
      Caption         =   "How to co&nnect:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1785
      Width           =   1500
   End
   Begin VB.Label lblServer
      Caption         =   "Game &server:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1335
      Width           =   1500
   End
   Begin VB.Label lblGame
      Caption         =   "&Game:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   885
      Width           =   1500
   End
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4065
      TabIndex        =   13
      Top             =   5910
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "&Save Details"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   5910
      Width           =   1455
   End
   Begin VB.CommandButton cmdIn
      Caption         =   "Count &Me In"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5910
      Width           =   1455
   End
   Begin VB.CheckBox chkRemind
      Caption         =   "&Remind me 5 minutes before"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5460
      Width           =   4920
   End
   Begin VB.ListBox lstIn
      Height          =   1230
      Left            =   240
      TabIndex        =   9
      Top             =   4140
      Width           =   4920
   End
   Begin VB.TextBox txtNote
      Height          =   1335
      Left            =   240
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2430
      Width           =   4920
   End
   Begin VB.PictureBox picIcon
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   240
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblIn
      Caption         =   "&Who's in:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3900
      Width           =   4920
   End
   Begin VB.Label lblDetails
      Caption         =   "&Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2190
      Width           =   4920
   End
   Begin VB.Label lblTime
      ForeColor       =   &H00636163&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   450
      Width           =   4560
   End
   Begin VB.Label lblName
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   210
      Width           =   4560
   End
End
Attribute VB_Name = "frmEventInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - What's On > Details: an event's details, who's in, and Remind Me.
Option Explicit

Private mId As String          ' the event's ID on the server
Private mFilling As Boolean    ' the boxes are being filled in, not typed in
Private mEdited As Boolean     ' the host has changed the details and not saved them yet
Private mGameKey() As String   ' the Game list's games: their keys, like quake3/quake3

Public Sub ShowEvent(ByVal id As String)
    mId = id
    mEdited = False
    RefreshInfo
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    Dim lines() As String, i As Long, p As Long
    lblName.FontBold = True
    cboGame.AddItem "(none)"
    ReDim mGameKey(0 To 0)
    lines = Split(LibGameList(), vbLf)
    For i = 0 To UBound(lines)
        p = InStr(lines(i), vbTab)
        If p > 1 Then
            cboGame.AddItem VBA.Mid$(lines(i), p + 1)
            ReDim Preserve mGameKey(0 To cboGame.ListCount - 1)
            mGameKey(cboGame.ListCount - 1) = Left$(lines(i), p - 1)
        End If
    Next
    If frmMain.HostMode() Then
        txtNote.Locked = False
        cmdSave.Visible = True
        cmdSave.Enabled = False
        lblDetails.Caption = "&Details (everyone can see these):"
    Else
        txtNote.Locked = True
        txtNote.BackColor = vbButtonFace
        txtServer.Locked = True
        txtServer.BackColor = vbButtonFace
        txtConnect.Locked = True
        txtConnect.BackColor = vbButtonFace
        cboGame.Enabled = False
        cmdSave.Visible = False
    End If
    ' older Palms Servers have no game for an event
    If Not frmMain.ServerHasProjector() Then
        lblGame.Visible = False
        cboGame.Visible = False
        lblServer.Visible = False
        txtServer.Visible = False
        lblConnect.Visible = False
        txtConnect.Visible = False
        lblDetails.Top = lblDetails.Top - 1350
        txtNote.Top = txtNote.Top - 1350
        lblIn.Top = lblIn.Top - 1350
        lstIn.Top = lstIn.Top - 1350
        chkRemind.Top = chkRemind.Top - 1350
        cmdIn.Top = cmdIn.Top - 1350
        cmdSave.Top = cmdSave.Top - 1350
        cmdClose.Top = cmdClose.Top - 1350
        Me.Height = Me.Height - 1350
    End If
End Sub

' The game's key in the list: picked, or added when the Cartridge Server no longer has it
Private Sub ShowGame(ByVal key As String)
    Dim i As Long
    If Len(key) = 0 Then
        cboGame.ListIndex = 0
        Exit Sub
    End If
    For i = 1 To cboGame.ListCount - 1
        If mGameKey(i) = key Then
            cboGame.ListIndex = i
            Exit Sub
        End If
    Next
    cboGame.AddItem key & " (not on the Cartridge Server)"
    ReDim Preserve mGameKey(0 To cboGame.ListCount - 1)
    mGameKey(cboGame.ListCount - 1) = key
    cboGame.ListIndex = cboGame.ListCount - 1
End Sub

' Called again by the main window whenever the schedule changes
Public Sub RefreshInfo()
    Dim i As Long, who() As String, k As Long, n As Long, nt As String
    i = frmMain.SchedIndexOf(mId)
    If i < 0 Then
        lblName.Caption = "Event removed"
        lblTime.Caption = "This event is no longer on the schedule."
        cmdIn.Enabled = False
        chkRemind.Enabled = False
        cmdSave.Enabled = False
        Exit Sub
    End If
    mFilling = True
    lblName.Caption = frmMain.SchedText(i, "name")
    lblTime.Caption = frmMain.SchedText(i, "when")
    picIcon.Cls
    DrawIcon16 picIcon, frmMain.SchedIconOf(i), 0, 0
    nt = frmMain.SchedText(i, "note")
    If Not mEdited Then
        ShowGame frmMain.SchedText(i, "game")
        txtServer.Text = frmMain.SchedText(i, "server")
        txtConnect.Text = frmMain.SchedText(i, "connect")
    End If
    If Not mEdited Then
        If Len(nt) = 0 And Not frmMain.HostMode() Then
            txtNote.Text = "No details have been added."
        Else
            txtNote.Text = NoteToText(nt)
        End If
    End If
    lstIn.Clear
    who = Split(frmMain.SchedText(i, "in"), ",")
    For k = 0 To UBound(who)
        If Len(who(k)) > 0 Then
            lstIn.AddItem who(k)
            n = n + 1
        End If
    Next
    lblIn.Caption = "&Who's in (" & n & "):"
    If frmMain.IsInEvent(i) Then
        cmdIn.Caption = "Count &Me Out"
    Else
        cmdIn.Caption = "Count &Me In"
    End If
    If frmMain.IsRemind(i) Then
        chkRemind.Value = vbChecked
    Else
        chkRemind.Value = vbUnchecked
    End If
    mFilling = False
End Sub

Private Sub txtNote_Change()
    If mFilling Then Exit Sub
    If Not frmMain.HostMode() Then Exit Sub
    mEdited = True
    cmdSave.Enabled = True
End Sub

Private Sub cboGame_Click()
    Edited
End Sub

Private Sub txtServer_Change()
    Edited
End Sub

Private Sub txtConnect_Change()
    Edited
End Sub

Private Sub Edited()
    If mFilling Then Exit Sub
    If Not frmMain.HostMode() Then Exit Sub
    mEdited = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    SaveNote
End Sub

Private Sub SaveNote()
    frmMain.SendLine "SCHEDNOTE|" & mId & "|" & NoteFromText(txtNote.Text)
    If cboGame.Visible Then
        frmMain.SendLine "SCHEDGAME|" & mId & "|" & mGameKey(cboGame.ListIndex) & "|" & NoBar(txtServer.Text) & "|" & NoBar(txtConnect.Text)
    End If
    mEdited = False
    cmdSave.Enabled = False
End Sub

Private Function NoBar(ByVal s As String) As String
    NoBar = Trim$(Replace(Replace(Replace(s, "|", "/"), vbCr, " "), vbLf, " "))
End Function

Private Sub cmdIn_Click()
    frmMain.ToggleIn frmMain.SchedIndexOf(mId)
End Sub

Private Sub chkRemind_Click()
    If mFilling Then Exit Sub
    frmMain.ToggleRemind frmMain.SchedIndexOf(mId)
End Sub

Private Sub cmdClose_Click()
    Dim r As VbMsgBoxResult
    If mEdited Then
        r = MsgBox("Save the changes to the details?", vbYesNoCancel + vbQuestion, "Event Details")
        If r = vbCancel Then Exit Sub
        If r = vbYes Then SaveNote
    End If
    mEdited = False
    Unload Me
End Sub
