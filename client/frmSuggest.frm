VERSION 5.00
Begin VB.Form frmSuggest
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suggest a Game"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmSuggest"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4065
      TabIndex        =   6
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove
      Caption         =   "&Remove"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Remove the selected game from the suggestions"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "&Suggest"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   3075
      Width           =   1260
   End
   Begin VB.ComboBox cboGame
      Height          =   315
      Left            =   990
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3105
      Width           =   2775
   End
   Begin VB.ListBox lstSugg
      Height          =   2085
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   870
      Width           =   4920
   End
   Begin VB.Label lblNew
      Caption         =   "&Game:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3150
      Width           =   735
   End
   Begin VB.Label lblPrompt
      Caption         =   "Games suggested for the next vote. Tick a game to back it (up to 3). The host picks the choices for the vote."
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4920
   End
End
Attribute VB_Name = "frmSuggest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - What's On > Suggest a Game.
Option Explicit

Private Const MAX_BACKED As Long = 3   ' the same limit as the server
Private mName() As String              ' the game on each row of the list
Private mFilling As Boolean            ' the list is being filled in: ticks aren't the user's
Private mPrompt As String

Private Sub Form_Load()
    Dim i As Long
    mFilling = False
    mPrompt = lblPrompt.Caption
    ReDim mName(0 To 0)
    ' the Game Library's games, to save typing
    For i = 0 To gShowCount - 1
        If Not InCombo(LibName(gShowId(i))) Then cboGame.AddItem LibName(gShowId(i))
    Next
    cmdRemove.Visible = frmMain.HostMode()
    RefreshList
End Sub

Private Function InCombo(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then
        InCombo = True
        Exit Function
    End If
    For i = 0 To cboGame.ListCount - 1
        If StrComp(cboGame.List(i), s, vbTextCompare) = 0 Then InCombo = True
    Next
End Function

' Called again by the main window whenever the list changes
Public Sub RefreshList()
    Dim i As Long, n As Long, sel As String, s As String
    If lstSugg.ListIndex >= 0 And lstSugg.ListIndex <= UBound(mName) Then sel = mName(lstSugg.ListIndex)
    mFilling = True
    lstSugg.Clear
    n = frmMain.SuggCount()
    If n > 0 Then
        ReDim mName(0 To n - 1)
    Else
        ReDim mName(0 To 0)
    End If
    For i = 0 To n - 1
        mName(i) = frmMain.SuggName(i)
        s = mName(i) & "   (" & frmMain.SuggVotes(i) & ": " & Replace(frmMain.SuggBy(i), ",", ", ") & ")"
        lstSugg.AddItem s
        If frmMain.SuggMine(i) Then lstSugg.Selected(lstSugg.NewIndex) = True
        If Len(sel) > 0 And StrComp(mName(i), sel, vbTextCompare) = 0 Then lstSugg.ListIndex = lstSugg.NewIndex
    Next
    mFilling = False
    cmdRemove.Enabled = (n > 0)
End Sub

' A message from the server (too many suggestions, the list is full)
Public Sub ShowError(ByVal msg As String)
    SayError msg
    RefreshList
End Sub

Private Sub SayError(ByVal msg As String)
    lblPrompt.ForeColor = vbRed
    lblPrompt.Caption = msg
    Beep
End Sub

Private Sub ResetPrompt()
    lblPrompt.ForeColor = vbButtonText
    lblPrompt.Caption = mPrompt
End Sub

Private Sub lstSugg_ItemCheck(Item As Integer)
    If mFilling Then Exit Sub
    If Item < 0 Or Item >= lstSugg.ListCount Or Item > UBound(mName) Then Exit Sub
    If lstSugg.Selected(Item) Then
        If frmMain.SuggMineCount() >= MAX_BACKED Then
            SayError "You can back up to " & MAX_BACKED & " games. Clear the tick on one first."
            mFilling = True
            lstSugg.Selected(Item) = False
            mFilling = False
            Exit Sub
        End If
        frmMain.SendLine "SUGGEST|" & mName(Item)
    Else
        frmMain.SendLine "UNSUGGEST|" & mName(Item)
    End If
    ResetPrompt
End Sub

Private Sub cmdAdd_Click()
    Dim s As String, i As Long
    s = Trim$(Replace(Replace(cboGame.Text, "|", "/"), ",", " "))
    If Len(s) = 0 Then
        SayError "Enter the name of a game, or select one from the list."
        cboGame.SetFocus
        Exit Sub
    End If
    If Len(s) > 40 Then s = Left$(s, 40)
    ' already backing it: nothing to do
    For i = 0 To frmMain.SuggCount() - 1
        If StrComp(frmMain.SuggName(i), s, vbTextCompare) = 0 And frmMain.SuggMine(i) Then
            cboGame.Text = ""
            Exit Sub
        End If
    Next
    If frmMain.SuggMineCount() >= MAX_BACKED Then
        SayError "You can back up to " & MAX_BACKED & " games. Clear the tick on one first."
        Exit Sub
    End If
    ResetPrompt
    frmMain.SendLine "SUGGEST|" & s
    cboGame.Text = ""
    cboGame.SetFocus
End Sub

' The host: take a game off the list
Private Sub cmdRemove_Click()
    Dim i As Long
    i = lstSugg.ListIndex
    If i < 0 Or i > UBound(mName) Or lstSugg.ListCount = 0 Then
        SayError "Select a game in the list first."
        Exit Sub
    End If
    If MsgBox("Remove " & mName(i) & " from the suggestions?", vbQuestion + vbOKCancel, "Suggest a Game") = vbOK Then
        frmMain.SendLine "SUGGDEL|" & mName(i)
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
