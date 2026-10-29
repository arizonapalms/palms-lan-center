VERSION 5.00
Begin VB.Form frmCupPlayers
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Players"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmCupPlayers"
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
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdSwap
      Caption         =   "S&wap"
      Height          =   375
      Left            =   1455
      TabIndex        =   6
      ToolTipText     =   "Swap the two selected players in the bracket"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove
      Caption         =   "&Remove"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Take the selected players off the list"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "A&dd"
      Default         =   -1  'True
      Height          =   375
      Left            =   4065
      TabIndex        =   4
      Top             =   3570
      Width           =   1095
   End
   Begin VB.ComboBox cboAdd
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Width           =   2865
   End
   Begin VB.ListBox lstPlayers
      Height          =   2595
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   840
      Width           =   4920
   End
   Begin VB.Label lblAdd
      Caption         =   "&Player:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3645
      Width           =   795
   End
   Begin VB.Label lblPrompt
      Caption         =   "Players"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4920
   End
End
Attribute VB_Name = "frmCupPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Tournaments > Players (host): add, remove and swap players.
Option Explicit

Private mCupId As String
Private mName() As String             ' the player on each row ("" for a bye)

Public Sub ShowCup(ByVal id As String)
    mCupId = id
    ReDim mName(0 To 0)
    RefreshCup
    If CupIndex(id) >= 0 Then Me.Show vbModal, frmMain
End Sub

' Called again by the main window whenever the tournament changes
Public Sub RefreshCup()
    Dim k As Long, p() As String, i As Long, n As Long, st As String, ppl() As String, sel As String
    k = CupIndex(mCupId)
    If k < 0 Then
        ' deleted while this was open
        If Me.Visible Then Unload Me
        Exit Sub
    End If
    st = gCupState(k)
    ' keep the selection where we can
    For i = 0 To lstPlayers.ListCount - 1
        If lstPlayers.Selected(i) And i <= UBound(mName) Then sel = sel & "," & LCase$(mName(i))
    Next
    sel = sel & ","
    lstPlayers.Clear
    n = 0
    If Len(gCupPlayers(k)) > 0 Then
        p = Split(gCupPlayers(k), ",")
        ReDim mName(0 To UBound(p))
        For i = 0 To UBound(p)
            If st = "signup" Then
                lstPlayers.AddItem p(i)
                mName(i) = p(i)
            ElseIf p(i) = "-" Then
                lstPlayers.AddItem "Match " & (i \ 2 + 1) & ":  (bye)"
                mName(i) = ""
            Else
                lstPlayers.AddItem "Match " & (i \ 2 + 1) & ":  " & p(i)
                mName(i) = p(i)
            End If
            If Len(mName(i)) > 0 And InStr(sel, "," & LCase$(mName(i)) & ",") > 0 Then lstPlayers.Selected(lstPlayers.NewIndex) = True
            If Len(mName(i)) > 0 Then n = n + 1
        Next
    Else
        ReDim mName(0 To 0)
    End If
    lblPrompt.ForeColor = vbButtonText
    Select Case st
    Case "signup"
        Me.Caption = "Players - " & gCupName(k)
        lblPrompt.Caption = n & " signed up. To add a player, type or select a nickname below. To remove players, select them and click Remove."
    Case "drawn"
        Me.Caption = "Bracket - " & gCupName(k)
        lblPrompt.Caption = "The first round, as drawn. To swap two players, select both and click Swap. To add or remove players, click Reopen Sign-ups on the toolbar."
    Case Else
        Me.Caption = "Players - " & gCupName(k)
        lblPrompt.Caption = "The tournament has started. The players can no longer be changed."
    End Select
    lblAdd.Enabled = (st = "signup")
    cboAdd.Enabled = (st = "signup")
    cmdAdd.Enabled = (st = "signup")
    cmdRemove.Enabled = (st = "signup")
    cmdSwap.Enabled = (st = "drawn")
    ' people signed in who aren't on the list yet
    cboAdd.Clear
    If st = "signup" Then
        ppl = Split(frmMain.PeopleHere(), "|")
        For i = 0 To UBound(ppl)
            If Len(ppl(i)) > 0 And Not CupIsIn(k, ppl(i)) Then cboAdd.AddItem ppl(i)
        Next
    End If
End Sub

' A message from the server (the tournament is full, sign-ups closed...)
Public Sub ShowError(ByVal msg As String)
    lblPrompt.ForeColor = vbRed
    lblPrompt.Caption = msg
End Sub

Private Function Picked() As String
    Dim i As Long, s As String
    For i = 0 To lstPlayers.ListCount - 1
        If lstPlayers.Selected(i) And i <= UBound(mName) Then
            If Len(mName(i)) > 0 Then s = s & "|" & mName(i)
        End If
    Next
    If Len(s) > 0 Then s = Mid$(s, 2)
    Picked = s
End Function

Private Sub cmdAdd_Click()
    Dim s As String, i As Long
    s = Trim$(cboAdd.Text)
    If Len(s) < 2 Or Len(s) > 12 Then
        MsgBox "Enter a nickname (2 to 12 characters).", vbExclamation, Me.Caption
        cboAdd.SetFocus
        Exit Sub
    End If
    For i = 1 To Len(s)
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_", Mid$(s, i, 1)) = 0 Then
            MsgBox "Nicknames can contain letters, numbers, hyphens (-) and underscores (_).", vbExclamation, Me.Caption
            cboAdd.SetFocus
            Exit Sub
        End If
    Next
    frmMain.SendLine "CUPADD|" & mCupId & "|" & s
    cboAdd.Text = ""
    cboAdd.SetFocus
End Sub

Private Sub cmdRemove_Click()
    Dim s As String, p() As String, i As Long
    s = Picked()
    If Len(s) = 0 Then
        MsgBox "Select the players to remove first.", vbExclamation, Me.Caption
        Exit Sub
    End If
    p = Split(s, "|")
    If UBound(p) = 0 Then
        If MsgBox("Remove " & p(0) & " from the tournament?", vbQuestion + vbOKCancel, Me.Caption) <> vbOK Then Exit Sub
    Else
        If MsgBox("Remove " & (UBound(p) + 1) & " players from the tournament?", vbQuestion + vbOKCancel, Me.Caption) <> vbOK Then Exit Sub
    End If
    For i = 0 To UBound(p)
        frmMain.SendLine "CUPREMOVE|" & mCupId & "|" & p(i)
    Next
End Sub

Private Sub cmdSwap_Click()
    Dim p() As String
    p = Split(Picked(), "|")
    If UBound(p) <> 1 Then
        MsgBox "Select two players to swap (click each one).", vbExclamation, Me.Caption
        Exit Sub
    End If
    frmMain.SendLine "CUPSWAP|" & mCupId & "|" & p(0) & "|" & p(1)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
