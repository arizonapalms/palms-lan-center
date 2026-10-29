VERSION 5.00
Begin VB.Form frmEventInfo
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Event Details"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmEventInfo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4065
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "&Save Details"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdIn
      Caption         =   "Count &Me In"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox chkRemind
      Caption         =   "&Remind me 5 minutes before"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4110
      Width           =   4920
   End
   Begin VB.ListBox lstIn
      Height          =   1230
      Left            =   240
      TabIndex        =   3
      Top             =   2790
      Width           =   4920
   End
   Begin VB.TextBox txtNote
      Height          =   1335
      Left            =   240
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblIn
      Caption         =   "&Who's in:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2550
      Width           =   4920
   End
   Begin VB.Label lblDetails
      Caption         =   "&Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4920
   End
   Begin VB.Label lblTime
      ForeColor       =   &H00636163&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   450
      Width           =   4560
   End
   Begin VB.Label lblName
      Height          =   255
      Left            =   600
      TabIndex        =   9
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

Public Sub ShowEvent(ByVal id As String)
    mId = id
    mEdited = False
    RefreshInfo
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    lblName.FontBold = True
    If frmMain.HostMode() Then
        txtNote.Locked = False
        cmdSave.Visible = True
        cmdSave.Enabled = False
        lblDetails.Caption = "&Details (everyone can see these):"
    Else
        txtNote.Locked = True
        txtNote.BackColor = vbButtonFace
        cmdSave.Visible = False
    End If
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

Private Sub cmdSave_Click()
    SaveNote
End Sub

Private Sub SaveNote()
    frmMain.SendLine "SCHEDNOTE|" & mId & "|" & NoteFromText(txtNote.Text)
    mEdited = False
    cmdSave.Enabled = False
End Sub

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
