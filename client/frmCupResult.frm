VERSION 5.00
Begin VB.Form frmCupResult
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Result"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "frmCupResult"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3465
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear
      Caption         =   "C&lear Result"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtScore
      Height          =   315
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1785
      Width           =   1200
   End
   Begin VB.OptionButton optB
      Caption         =   "B won"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1350
      Width           =   3960
   End
   Begin VB.OptionButton optA
      Caption         =   "A won"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1020
      Width           =   3960
   End
   Begin VB.Label lblScoreHint
      Caption         =   "Optional. For example, 16-12."
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1830
      Width           =   2160
   End
   Begin VB.Label lblScore
      Caption         =   "&Score:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1830
      Width           =   735
   End
   Begin VB.Label lblInfo
      Caption         =   "Select the winner."
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4320
   End
   Begin VB.Label lblMatch
      Caption         =   "Final: A vs B"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4320
   End
End
Attribute VB_Name = "frmCupResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Tournaments > Report Result, and Enter or Confirm Result (host).
Option Explicit

Private mCupId As String
Private mRound As Long
Private mMatch As Long
Private mHost As Boolean

Public Sub ShowMatch(ByVal id As String, ByVal r As Long, ByVal i As Long)
    Dim k As Long, a As String, b As String, win As String, score As String, rby As String, rw As String, rs As String
    k = CupIndex(id)
    If k < 0 Then
        Unload Me
        Exit Sub
    End If
    mCupId = id
    mRound = r
    mMatch = i
    mHost = frmMain.HostMode()
    a = CupM(k, r, i, "a")
    b = CupM(k, r, i, "b")
    win = CupM(k, r, i, "win")
    score = CupM(k, r, i, "score")
    rby = CupM(k, r, i, "repby")
    rw = CupM(k, r, i, "repwin")
    rs = CupM(k, r, i, "repscore")
    lblMatch.FontBold = True
    lblMatch.Caption = gCupName(k) & ", " & CupRoundWord(k, r) & ": " & a & " vs " & b
    optA.Caption = a & " won"
    optB.Caption = b & " won"
    cmdClear.Visible = mHost And Len(win) > 0
    If mHost Then
        If Len(win) > 0 Then
            Me.Caption = "Change Result"
            cmdOK.Caption = "OK"
            lblInfo.Caption = WonWords(a, b, win, score) & " A result can be changed until the next match has one."
            Pick win, score
        ElseIf Len(rby) > 0 Then
            Me.Caption = "Confirm Result"
            cmdOK.Caption = "&Confirm"
            lblInfo.Caption = "Reported by " & rby & ": " & WonWords(a, b, rw, rs) & " Click Confirm, or correct the result first."
            Pick rw, rs
        Else
            Me.Caption = "Enter Result"
            cmdOK.Caption = "OK"
            lblInfo.Caption = "Select the winner. The result goes straight into the bracket."
        End If
    Else
        Me.Caption = "Report Result"
        cmdOK.Caption = "&Report"
        If Len(rby) > 0 Then
            lblInfo.Caption = "Reported by " & rby & ": " & WonWords(a, b, rw, rs) & " Reporting again replaces it."
            Pick rw, rs
        Else
            lblInfo.Caption = "Select the winner. The host confirms the result before it goes into the bracket."
        End If
    End If
    Me.Show vbModal, frmMain
End Sub

Private Sub Pick(ByVal side As String, ByVal score As String)
    optA.Value = (side = "a")
    optB.Value = (side = "b")
    txtScore.Text = score
End Sub

Private Function WonWords(ByVal a As String, ByVal b As String, ByVal side As String, ByVal score As String) As String
    Dim w As String
    If side = "a" Then w = a Else w = b
    If Len(score) > 0 Then
        WonWords = w & " won " & score & "."
    Else
        WonWords = w & " won."
    End If
End Function

Private Sub cmdOK_Click()
    Dim side As String, score As String
    If optA.Value Then
        side = "a"
    ElseIf optB.Value Then
        side = "b"
    Else
        MsgBox "Select the winner.", vbExclamation, Me.Caption
        optA.SetFocus
        Exit Sub
    End If
    score = Trim$(Replace(Replace(txtScore.Text, "|", "/"), "~", "-"))
    If mHost Then
        frmMain.SendLine "CUPRESULT|" & mCupId & "|" & mRound & "|" & mMatch & "|" & side & "|" & score
    Else
        frmMain.SendLine "CUPREPORT|" & mCupId & "|" & mRound & "|" & mMatch & "|" & side & "|" & score
        frmMain.CupNote "Result reported. The host confirms it."
    End If
    Unload Me
End Sub

' The host: take a result out again (only while the next match has no result)
Private Sub cmdClear_Click()
    If MsgBox("Clear the result of this match?", vbQuestion + vbOKCancel, Me.Caption) <> vbOK Then Exit Sub
    frmMain.SendLine "CUPRESULT|" & mCupId & "|" & mRound & "|" & mMatch & "||"
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
