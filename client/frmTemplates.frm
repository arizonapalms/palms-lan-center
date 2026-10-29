VERSION 5.00
Begin VB.Form frmTemplates
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule Templates"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "frmTemplates"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdUse
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.OptionButton optAdd
      Caption         =   "Add to t&oday's schedule"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   5220
      Width           =   2415
   End
   Begin VB.OptionButton optReplace
      Caption         =   "&Replace today's schedule"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5220
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "&Save Template"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtLines
      Height          =   2895
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   2760
      MaxLength       =   40
      TabIndex        =   5
      Top             =   420
      Width           =   4935
   End
   Begin VB.CommandButton cmdDelete
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew
      Caption         =   "&New"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ListBox lstTpl
      Height          =   3570
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   2295
   End
   Begin VB.Label lblStatus
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   4620
      UseMnemonic     =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblHint
      Caption         =   "Enter times as 10:00 AM or 19:30. Icons are chosen from keywords (lunch, game, final). To set an icon, start the text with [trophy], [pizza] and so on."
      ForeColor       =   &H00636163&
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   4020
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblLines
      Caption         =   "&Events (one per line, starting with the time):"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblName
      Caption         =   "N&ame:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
   Begin VB.Label lblList
      Caption         =   "&Templates:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Schedule templates (host).
Option Explicit

Private mIds() As String
Private mCount As Long
Private mBuilding As Boolean
Private mCurId As String       ' the template on show; "new" for one not saved yet
Private mWantId As String      ' pick this one when the list comes back

Private Sub Form_Load()
    ReDim mIds(0 To 0)
    lblStatus.Caption = "Retrieving templates..."
    frmMain.SendLine "TPLLIST"
End Sub

' frmMain passes on everything the server says about templates
Public Sub OnServer(ByVal lineText As String)
    Dim f() As String, i As Long, pick As Long
    f = Split(lineText, "|")
    Select Case f(0)
    Case "TPL"
        If Not mBuilding Then
            mBuilding = True
            lstTpl.Clear
            mCount = 0
        End If
        ReDim Preserve mIds(0 To mCount)
        mIds(mCount) = Part(f, 1)
        lstTpl.AddItem Rest(f, 2)
        mCount = mCount + 1
    Case "TPLEND"
        If Not mBuilding Then
            lstTpl.Clear
            mCount = 0
        End If
        mBuilding = False
        If lblStatus.Caption = "Retrieving templates..." Then lblStatus.Caption = ""
        pick = -1
        For i = 0 To mCount - 1
            If mIds(i) = mWantId Then pick = i
        Next
        If pick = -1 And mCount > 0 And mCurId <> "new" Then pick = 0
        If pick >= 0 Then lstTpl.ListIndex = pick
    Case "TPLTEXT"
        If Part(f, 1) <> mCurId Then Exit Sub
        txtName.Text = Part(f, 2)
        txtLines.Text = Replace(Rest(f, 3), "~", vbCrLf)
        If lblStatus.Caption = "Loading..." Then lblStatus.Caption = ""
    Case "TPLSAVED"
        mWantId = Part(f, 1)
        mCurId = mWantId
        lblStatus.Caption = "Saved."
    Case "TPLAPPLIED"
        If Val(Part(f, 2)) > 0 Then
            MsgBox "Today's schedule updated (" & Part(f, 1) & " events). " & Part(f, 2) & " lines did not start with a time and were skipped.", vbInformation, "Schedule Templates"
        Else
            MsgBox "Today's schedule updated (" & Part(f, 1) & " events).", vbInformation, "Schedule Templates"
        End If
        Unload Me
    Case "ERR"
        lblStatus.Caption = Rest(f, 2)
        Beep
    End Select
End Sub

Private Sub lstTpl_Click()
    If lstTpl.ListIndex < 0 Or lstTpl.ListIndex >= mCount Then Exit Sub
    If mIds(lstTpl.ListIndex) = mCurId And Len(txtLines.Text) > 0 Then Exit Sub
    mCurId = mIds(lstTpl.ListIndex)
    txtName.Text = lstTpl.List(lstTpl.ListIndex)
    txtLines.Text = ""
    lblStatus.Caption = "Loading..."
    frmMain.SendLine "TPLGET|" & mCurId
End Sub

Private Sub cmdNew_Click()
    lstTpl.ListIndex = -1
    mCurId = "new"
    txtName.Text = "Untitled"
    txtLines.Text = "10:00 AM  Doors open" & vbCrLf
    lblStatus.Caption = "Enter the times and events, then click Save Template."
    On Error Resume Next
    txtName.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If mCurId = "" Or mCurId = "new" Then Exit Sub
    If MsgBox("Delete template " & txtName.Text & "?", vbQuestion + vbOKCancel, "Schedule Templates") <> vbOK Then Exit Sub
    frmMain.SendLine "TPLDEL|" & mCurId
    mCurId = ""
    mWantId = ""
    txtName.Text = ""
    txtLines.Text = ""
End Sub

Private Sub cmdSave_Click()
    Dim lines As String, nm As String
    lines = Encoded()
    If Len(lines) = 0 Then
        MsgBox "Enter at least one event (for example, 12:15 PM Game 1).", vbExclamation, "Schedule Templates"
        Exit Sub
    End If
    nm = Trim$(Replace(Replace(txtName.Text, "|", "-"), "~", "-"))
    If Len(nm) = 0 Then nm = "Untitled"
    If Len(mCurId) = 0 Then mCurId = "new"
    lblStatus.Caption = "Saving..."
    frmMain.SendLine "TPLSAVE|" & mCurId & "|" & nm & "|" & lines
End Sub

Private Sub cmdUse_Click()
    Dim lines As String, mode As String
    lines = Encoded()
    If Len(lines) = 0 Then
        MsgBox "Enter at least one event (for example, 12:15 PM Game 1).", vbExclamation, "Schedule Templates"
        Exit Sub
    End If
    If optAdd.Value Then
        mode = "add"
    Else
        mode = "replace"
        If MsgBox("Replace today's schedule with this template? The change is shown on every PC immediately.", vbQuestion + vbOKCancel, "Schedule Templates") <> vbOK Then Exit Sub
    End If
    lblStatus.Caption = "Updating today's schedule..."
    frmMain.SendLine "TPLAPPLY|" & mode & "|" & lines
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' The lines as one piece for the server: "~" between lines, no "|" or "~" inside them
Private Function Encoded() As String
    Dim parts() As String, i As Long, s As String, out As String
    parts = Split(Replace(txtLines.Text, vbCr, ""), vbLf)
    For i = 0 To UBound(parts)
        s = Trim$(Replace(Replace(parts(i), "~", "-"), "|", "-"))
        If Len(s) > 0 Then
            If Len(out) > 0 Then out = out & "~"
            out = out & s
        End If
    Next
    Encoded = out
End Function

Private Function Part(f() As String, ByVal i As Long) As String
    If i <= UBound(f) Then Part = f(i)
End Function

Private Function Rest(f() As String, ByVal i As Long) As String
    Dim k As Long, s As String
    If i > UBound(f) Then Exit Function
    s = f(i)
    For k = i + 1 To UBound(f)
        s = s & "|" & f(k)
    Next
    Rest = s
End Function
