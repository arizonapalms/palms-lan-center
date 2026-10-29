VERSION 5.00
Begin VB.Form frmCupEdit
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Tournament"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmCupEdit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4065
      TabIndex        =   13
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   3780
      Width           =   1095
   End
   Begin VB.TextBox txtNotes
      Height          =   1215
      Left            =   1560
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1830
      Width           =   3600
   End
   Begin VB.ComboBox cboMax
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1425
      Width           =   1200
   End
   Begin VB.TextBox txtTime
      Height          =   315
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   5
      Top             =   1020
      Width           =   1200
   End
   Begin VB.ComboBox cboGame
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   615
      Width           =   3600
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   210
      Width           =   3600
   End
   Begin VB.Label lblNotesHint
      Caption         =   "Rules, the server address, maps. Shown to everyone on the Tournaments page."
      Height          =   435
      Left            =   1560
      TabIndex        =   11
      Top             =   3105
      Width           =   3600
   End
   Begin VB.Label lblNotes
      Caption         =   "&Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label lblMaxHint
      Caption         =   "Most players who can sign up."
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   1470
      Width           =   2280
   End
   Begin VB.Label lblMax
      Caption         =   "&Players:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1470
      Width           =   1215
   End
   Begin VB.Label lblTimeHint
      Caption         =   "Optional. For example, 8:30 PM."
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   1065
      Width           =   2280
   End
   Begin VB.Label lblTime
      Caption         =   "&Starts at:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1065
      Width           =   1215
   End
   Begin VB.Label lblGame
      Caption         =   "&Game:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblName
      Caption         =   "&Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "frmCupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Tournaments > New Tournament and Edit (host).
Option Explicit

Private mCupId As String               ' the tournament being changed, "" for a new one
Private mSigned As Long                ' players signed up already (the maximum can't go below it)

' id: the tournament to change, or "" for a new one
Public Sub EditCup(ByVal id As String)
    Dim k As Long, i As Long, most As Long
    mCupId = id
    k = CupIndex(id)
    If Len(id) > 0 And k < 0 Then
        Unload Me
        Exit Sub
    End If
    ' the Game Library's games, to save typing
    cboGame.Clear
    For i = 0 To gShowCount - 1
        If Not InCombo(LibName(gShowId(i))) Then cboGame.AddItem LibName(gShowId(i))
    Next
    most = 0
    If k >= 0 Then
        Me.Caption = "Edit Tournament"
        txtName.Text = gCupName(k)
        cboGame.Text = gCupGame(k)
        txtTime.Text = gCupTime(k)
        txtNotes.Text = Replace(gCupNotes(k), "~", vbCrLf)
        most = gCupMax(k)
        If gCupState(k) = "signup" Then mSigned = CupPlayerCount(k)
    Else
        Me.Caption = "New Tournament"
        mSigned = 0
    End If
    FillMax most
    Me.Show vbModal, frmMain
End Sub

Private Sub FillMax(ByVal most As Long)
    Dim v As Variant, i As Long
    cboMax.Clear
    cboMax.AddItem "No limit"
    cboMax.ItemData(cboMax.NewIndex) = 0
    For Each v In Array(4, 8, 16, 32, 64)
        cboMax.AddItem CStr(v)
        cboMax.ItemData(cboMax.NewIndex) = v
    Next
    ' a number set some other way (HELP on the server lets the host type any)
    If most > 0 Then
        For i = 0 To cboMax.ListCount - 1
            If cboMax.ItemData(i) = most Then
                cboMax.ListIndex = i
                Exit Sub
            End If
        Next
        cboMax.AddItem CStr(most)
        cboMax.ItemData(cboMax.NewIndex) = most
        cboMax.ListIndex = cboMax.NewIndex
    Else
        cboMax.ListIndex = 0
    End If
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

' One line of text, safe to send: no "|", and no line breaks
Private Function OneLine(ByVal s As String, ByVal most As Long) As String
    s = Replace(Replace(Replace(s, "|", "/"), vbCr, " "), vbLf, " ")
    s = Trim$(Replace(s, "~", "-"))
    If Len(s) > most Then s = Left$(s, most)
    OneLine = s
End Function

Private Sub cmdOK_Click()
    Dim nm As String, notes As String, most As Long, fields As String
    nm = OneLine(txtName.Text, 30)
    If Len(nm) = 0 Then
        MsgBox "Enter a name for the tournament (for example, Quake III Duel).", vbExclamation, Me.Caption
        txtName.SetFocus
        Exit Sub
    End If
    most = cboMax.ItemData(cboMax.ListIndex)
    If most > 0 And most < mSigned Then
        MsgBox mSigned & " players have already signed up. Select a higher number, or No limit.", vbExclamation, Me.Caption
        cboMax.SetFocus
        Exit Sub
    End If
    ' the details keep their lines: ~ between them
    notes = Replace(Replace(Replace(txtNotes.Text, "|", "/"), vbCrLf, "~"), vbLf, "~")
    notes = Trim$(Replace(notes, vbCr, ""))
    Do While InStr(notes, "~~") > 0
        notes = Replace(notes, "~~", "~")
    Loop
    If Len(notes) > 300 Then notes = Left$(notes, 300)
    fields = most & "|" & OneLine(txtTime.Text, 12) & "|" & OneLine(cboGame.Text, 40) & "|" & nm & "|" & notes
    If Len(mCupId) > 0 Then
        frmMain.SendLine "CUPEDIT|" & mCupId & "|" & fields
    Else
        frmMain.SendLine "CUPNEW|" & fields
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
