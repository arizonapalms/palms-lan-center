VERSION 5.00
Begin VB.Form frmIgnore
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mute Players"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "frmIgnore"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3660
      Width           =   1095
   End
   Begin VB.ListBox lstPeople
      Height          =   2535
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   660
      Width           =   4320
   End
   Begin VB.Label lblNote
      Caption         =   "Announcements from the host are always shown."
      ForeColor       =   &H00636163&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3300
      Width           =   4320
   End
   Begin VB.Label lblPrompt
      Caption         =   "&Messages from the selected users are hidden on this PC."
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4320
   End
End
Attribute VB_Name = "frmIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Chat > Mute Players: hides chat from chosen people, on this PC only.
Option Explicit

Private Sub Form_Load()
    Dim people() As String, ign() As String, i As Long
    people = Split(frmMain.OtherPeople(), "|")
    For i = 0 To UBound(people)
        If Len(people(i)) > 0 Then AddName people(i), IsIgnored(people(i))
    Next
    ' people ignored before who aren't signed in right now
    ign = Split(gIgnore, ";")
    For i = 0 To UBound(ign)
        If Len(ign(i)) > 0 Then
            If Not Listed(ign(i)) Then AddName ign(i), True
        End If
    Next
    If lstPeople.ListCount = 0 Then lblPrompt.Caption = "No other users are signed in."
    If lstPeople.ListCount > 0 Then lstPeople.ListIndex = 0
End Sub

Private Sub AddName(ByVal nm As String, ByVal isChecked As Boolean)
    lstPeople.AddItem nm
    lstPeople.Selected(lstPeople.NewIndex) = isChecked
End Sub

Private Function Listed(ByVal nm As String) As Boolean
    Dim i As Long
    For i = 0 To lstPeople.ListCount - 1
        If StrComp(lstPeople.List(i), nm, vbTextCompare) = 0 Then Listed = True
    Next
End Function

Private Sub cmdOK_Click()
    Dim i As Long, s As String
    For i = 0 To lstPeople.ListCount - 1
        If lstPeople.Selected(i) Then s = s & LCase$(lstPeople.List(i)) & ";"
    Next
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    gIgnore = s
    SaveSettings
    frmMain.IgnoreChanged
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
