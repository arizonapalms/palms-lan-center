VERSION 5.00
Begin VB.Form frmMapEdit
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seat Map"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "frmMapEdit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdUse
      Caption         =   "&Use This Venue"
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "&Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   6000
      Width           =   1095
   End
   Begin VB.PictureBox picPreview
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E7F7FF&
      Height          =   5235
      Left            =   4800
      ScaleMode       =   3  'Pixel
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   480
      Width           =   4575
   End
   Begin VB.Frame fraTable
      Caption         =   "Selected Table"
      Height          =   1335
      Left            =   240
      TabIndex        =   16
      Top             =   3900
      Width           =   4335
      Begin VB.CheckBox chkBoth
         Caption         =   "Seats on b&oth sides"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.ComboBox cboPer
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   645
         Width           =   975
      End
      Begin VB.TextBox txtPrefix
         Height          =   315
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   20
         Top             =   645
         Width           =   495
      End
      Begin VB.TextBox txtTName
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   18
         Top             =   255
         Width           =   2895
      End
      Begin VB.Label lblPer
         Caption         =   "Seats &per side:"
         Height          =   255
         Left            =   1980
         TabIndex        =   21
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblPrefix
         Caption         =   "Seat le&tter:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblTName
         Caption         =   "Na&me:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDown
      Caption         =   "Move Do&wn"
      Height          =   345
      Left            =   3360
      TabIndex        =   15
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdUp
      Caption         =   "Move &Up"
      Height          =   345
      Left            =   3360
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove
      Caption         =   "&Remove"
      Height          =   345
      Left            =   3360
      TabIndex        =   13
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd
      Caption         =   "&Add Table"
      Height          =   345
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstTables
      Height          =   1620
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox cboCols
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1455
      Width           =   735
   End
   Begin VB.TextBox txtStage
      Height          =   315
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   7
      ToolTipText     =   "Shown across the top of the map. Leave blank for none."
      Top             =   1065
      Width           =   2895
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   5
      Top             =   675
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete
      Caption         =   "&Delete"
      Height          =   345
      Left            =   3840
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton cmdNew
      Caption         =   "&New"
      Height          =   345
      Left            =   3000
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
   Begin VB.ComboBox cboVenue
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   195
      Width           =   1815
   End
   Begin VB.Label lblStatus
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   6060
      UseMnemonic     =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblHint
      Caption         =   "Seats are numbered along one side first (A01 to A04), then along the other side (A05 to A08)."
      ForeColor       =   &H00636163&
      Height          =   435
      Left            =   240
      TabIndex        =   29
      Top             =   5340
      UseMnemonic     =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblPreview
      Caption         =   "Preview:"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label lblTables
      Caption         =   "Ta&bles:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblCols
      Caption         =   "Tables side by s&ide:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Label lblStage
      Caption         =   "Stage or s&creen:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label lblName
      Caption         =   "Venue nam&e:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblVenue
      Caption         =   "Venue:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - The seat map editor (host).
Option Explicit

Private mMap As SeatMap           ' the venue being edited
Private mIds() As String          ' the venue ids, in the same order as cboVenue
Private mCount As Long
Private mActiveId As String       ' the venue in use
Private mWantId As String         ' pick this venue when the list comes back
Private mLoading As Boolean       ' filling the boxes in, so it doesn't count as a change
Private mDirty As Boolean         ' changed since it was loaded or saved
Private mListBusy As Boolean      ' the venue list is coming in
Private mLoadAfterList As Boolean ' fetch the picked venue once the list arrives
Private mUseAfterSave As Boolean  ' Use this venue was clicked with changes not saved
Private mUsing As Boolean         ' waiting for the server to switch everyone over

Private Sub Form_Load()
    Dim i As Long
    ReDim mIds(0 To 0)
    For i = 1 To 4
        cboCols.AddItem CStr(i)
    Next
    For i = 1 To 12
        cboPer.AddItem CStr(i)
    Next
    NewMap
    mDirty = False
    mLoadAfterList = True
    lblStatus.Caption = "Retrieving venues..."
    frmMain.SendLine "MAPLIST"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mDirty Then Exit Sub
    If MsgBox("Changes to " & mMap.Name & " have not been saved. Close without saving?", vbQuestion + vbYesNo, "Seat Map") <> vbYes Then Cancel = 1
End Sub

' ---- messages from the server (frmMain passes them on)

Public Sub OnServer(ByVal lineText As String)
    Dim f() As String, pick As Long
    f = Split(lineText, "|")
    Select Case f(0)
    Case "MAPL"
        If Not mListBusy Then
            mListBusy = True
            mLoading = True
            cboVenue.Clear
            mCount = 0
            mLoading = False
        End If
        ReDim Preserve mIds(0 To mCount)
        mIds(mCount) = Part(f, 1)
        If Part(f, 2) = "1" Then
            mActiveId = Part(f, 1)
            cboVenue.AddItem Rest(f, 3) & " (in use)"
        Else
            cboVenue.AddItem Rest(f, 3)
        End If
        mCount = mCount + 1
    Case "MAPLEND"
        If Not mListBusy Then
            cboVenue.Clear
            mCount = 0
        End If
        mListBusy = False
        pick = IdIndex(mWantId)
        If pick < 0 Then pick = IdIndex(mMap.Id)
        If pick < 0 And mMap.Id <> "new" Then pick = IdIndex(mActiveId)
        mLoading = True
        cboVenue.ListIndex = pick
        mLoading = False
        If mLoadAfterList And pick >= 0 Then
            mLoadAfterList = False
            LoadVenue mIds(pick)
        ElseIf lblStatus.Caption = "Retrieving venues..." Then
            lblStatus.Caption = ""
        End If
        If mUsing Then
            mUsing = False
            lblStatus.Caption = mMap.Name & " is now the active seat map."
        End If
    Case "MAPDEF"
        If Part(f, 1) <> mWantId Then Exit Sub
        MapFromFields f, 1, mMap
        ShowMap
        mDirty = False
        lblStatus.Caption = ""
    Case "MAPSAVED"
        mMap.Id = Part(f, 1)
        mWantId = mMap.Id
        mDirty = False
        If mUseAfterSave Then
            mUseAfterSave = False
            UseVenue
        Else
            lblStatus.Caption = "Saved."
        End If
    Case "ERR"
        mUseAfterSave = False
        mUsing = False
        lblStatus.Caption = Rest(f, 2)
        Beep
    End Select
End Sub

Private Function IdIndex(ByVal id As String) As Long
    Dim i As Long
    IdIndex = -1
    If Len(id) = 0 Then Exit Function
    For i = 0 To mCount - 1
        If mIds(i) = id Then IdIndex = i
    Next
End Function

Private Sub LoadVenue(ByVal id As String)
    mWantId = id
    lblStatus.Caption = "Loading..."
    frmMain.SendLine "MAPGET|" & id
End Sub

' ---- showing a map in the boxes

' A fresh venue: a big screen and one table of eight
Private Sub NewMap()
    mMap.Id = "new"
    mMap.Name = "New Venue"
    mMap.Stage = "Big screen and stage"
    mMap.Cols = 1
    mMap.Count = 1
    ReDim mMap.Tables(0 To 0)
    mMap.Tables(0).Name = "Table A"
    mMap.Tables(0).Prefix = "A"
    mMap.Tables(0).PerSide = 4
    mMap.Tables(0).Sides = 2
    ShowMap
End Sub

Private Sub ShowMap()
    mLoading = True
    txtName.Text = mMap.Name
    txtStage.Text = mMap.Stage
    cboCols.ListIndex = mMap.Cols - 1
    mLoading = False
    FillTables 0
    DrawPreview
End Sub

Private Function TableText(ByVal t As Long) As String
    Dim last As Long
    With mMap.Tables(t)
        last = .PerSide * .Sides
        TableText = .Name & "  (" & SeatLabel(.Prefix, 1) & " to " & SeatLabel(.Prefix, last) & ")"
    End With
End Function

Private Sub FillTables(ByVal sel As Long)
    Dim t As Long, was As Boolean
    was = mLoading
    mLoading = True
    lstTables.Clear
    For t = 0 To mMap.Count - 1
        lstTables.AddItem TableText(t)
    Next
    If sel >= mMap.Count Then sel = mMap.Count - 1
    If sel >= 0 Then lstTables.ListIndex = sel
    mLoading = was
    ShowTable
End Sub

Private Function SelTable() As Long
    SelTable = lstTables.ListIndex
    If SelTable >= mMap.Count Then SelTable = -1
End Function

Private Sub ShowTable()
    Dim i As Long, was As Boolean
    i = SelTable()
    was = mLoading
    mLoading = True
    fraTable.Enabled = (i >= 0)
    txtTName.Enabled = (i >= 0)
    txtPrefix.Enabled = (i >= 0)
    cboPer.Enabled = (i >= 0)
    chkBoth.Enabled = (i >= 0)
    If i >= 0 Then
        txtTName.Text = mMap.Tables(i).Name
        txtPrefix.Text = mMap.Tables(i).Prefix
        cboPer.ListIndex = mMap.Tables(i).PerSide - 1
        If mMap.Tables(i).Sides = 2 Then
            chkBoth.Value = vbChecked
        Else
            chkBoth.Value = vbUnchecked
        End If
    Else
        txtTName.Text = ""
        txtPrefix.Text = ""
    End If
    mLoading = was
End Sub

Private Sub DrawPreview()
    Dim lay As MapLayout
    picPreview.Cls
    LayoutMap mMap, picPreview.ScaleWidth, True, lay
    DrawMap picPreview, lay, 0
End Sub

' Something changed: remember it isn't saved, and redraw the preview
Private Sub Changed()
    mDirty = True
    lblStatus.Caption = "Not saved."
    DrawPreview
End Sub

Private Sub UpdateRow(ByVal i As Long)
    Dim was As Boolean
    was = mLoading
    mLoading = True
    lstTables.List(i) = TableText(i)
    mLoading = was
End Sub

' ---- the venue boxes

Private Sub cboVenue_Click()
    Dim back As Long
    If mLoading Then Exit Sub
    If cboVenue.ListIndex < 0 Then Exit Sub
    If mIds(cboVenue.ListIndex) = mMap.Id And Not mDirty Then Exit Sub
    If mDirty Then
        If MsgBox("Changes to " & mMap.Name & " have not been saved. Open another venue without saving?", vbQuestion + vbYesNo, "Seat Map") <> vbYes Then
            back = IdIndex(mMap.Id)
            mLoading = True
            cboVenue.ListIndex = back
            mLoading = False
            Exit Sub
        End If
    End If
    mDirty = False
    LoadVenue mIds(cboVenue.ListIndex)
End Sub

Private Sub txtName_Change()
    If mLoading Then Exit Sub
    mMap.Name = txtName.Text
    Changed
End Sub

Private Sub txtStage_Change()
    If mLoading Then Exit Sub
    mMap.Stage = txtStage.Text
    Changed
End Sub

Private Sub cboCols_Click()
    If mLoading Then Exit Sub
    mMap.Cols = cboCols.ListIndex + 1
    Changed
End Sub

' ---- the tables

Private Sub lstTables_Click()
    If mLoading Then Exit Sub
    ShowTable
End Sub

Private Sub txtTName_Change()
    Dim i As Long
    If mLoading Then Exit Sub
    i = SelTable()
    If i < 0 Then Exit Sub
    mMap.Tables(i).Name = txtTName.Text
    UpdateRow i
    Changed
End Sub

' Seat letters are one or two capital letters
Private Sub txtPrefix_Change()
    Dim i As Long, k As Long, s As String, t As String
    If mLoading Then Exit Sub
    i = SelTable()
    If i < 0 Then Exit Sub
    s = UCase$(txtPrefix.Text)
    For k = 1 To Len(s)
        If Mid$(s, k, 1) Like "[A-Z]" Then t = t & Mid$(s, k, 1)
    Next
    If t <> txtPrefix.Text Then
        mLoading = True
        txtPrefix.Text = t
        txtPrefix.SelStart = Len(t)
        mLoading = False
    End If
    mMap.Tables(i).Prefix = t
    UpdateRow i
    Changed
End Sub

Private Sub cboPer_Click()
    Dim i As Long
    If mLoading Then Exit Sub
    i = SelTable()
    If i < 0 Then Exit Sub
    mMap.Tables(i).PerSide = cboPer.ListIndex + 1
    UpdateRow i
    Changed
End Sub

Private Sub chkBoth_Click()
    Dim i As Long
    If mLoading Then Exit Sub
    i = SelTable()
    If i < 0 Then Exit Sub
    If chkBoth.Value = vbChecked Then
        mMap.Tables(i).Sides = 2
    Else
        mMap.Tables(i).Sides = 1
    End If
    UpdateRow i
    Changed
End Sub

' The first letter no table is using yet
Private Function FreeLetter() As String
    Dim c As Long, t As Long, used As Boolean
    For c = 0 To 25
        used = False
        For t = 0 To mMap.Count - 1
            If mMap.Tables(t).Prefix = Chr$(65 + c) Then used = True
        Next
        If Not used Then
            FreeLetter = Chr$(65 + c)
            Exit Function
        End If
    Next
    FreeLetter = "Z"
End Function

Private Sub cmdAdd_Click()
    Dim n As Long, letter As String
    If mMap.Count >= 26 Then
        lblStatus.Caption = "A seat map can have a maximum of 26 tables."
        Beep
        Exit Sub
    End If
    n = mMap.Count
    letter = FreeLetter()
    ReDim Preserve mMap.Tables(0 To n)
    mMap.Tables(n).Name = "Table " & letter
    mMap.Tables(n).Prefix = letter
    If n > 0 Then
        mMap.Tables(n).PerSide = mMap.Tables(n - 1).PerSide
        mMap.Tables(n).Sides = mMap.Tables(n - 1).Sides
    Else
        mMap.Tables(n).PerSide = 4
        mMap.Tables(n).Sides = 2
    End If
    mMap.Count = n + 1
    FillTables n
    Changed
    On Error Resume Next
    txtTName.SetFocus
End Sub

Private Sub cmdRemove_Click()
    Dim i As Long, k As Long
    i = SelTable()
    If i < 0 Then Exit Sub
    If mMap.Count <= 1 Then
        lblStatus.Caption = "A seat map must have at least one table."
        Beep
        Exit Sub
    End If
    For k = i To mMap.Count - 2
        mMap.Tables(k) = mMap.Tables(k + 1)
    Next
    mMap.Count = mMap.Count - 1
    ReDim Preserve mMap.Tables(0 To mMap.Count - 1)
    FillTables i
    Changed
End Sub

Private Sub cmdUp_Click()
    MoveTable -1
End Sub

Private Sub cmdDown_Click()
    MoveTable 1
End Sub

Private Sub MoveTable(ByVal stp As Long)
    Dim i As Long, j As Long, tmp As MapTable
    i = SelTable()
    j = i + stp
    If i < 0 Or j < 0 Or j >= mMap.Count Then Exit Sub
    tmp = mMap.Tables(i)
    mMap.Tables(i) = mMap.Tables(j)
    mMap.Tables(j) = tmp
    FillTables j
    Changed
End Sub

' ---- venues

Private Sub cmdNew_Click()
    If mDirty Then
        If MsgBox("Changes to " & mMap.Name & " have not been saved. Create a new venue without saving?", vbQuestion + vbYesNo, "Seat Map") <> vbYes Then Exit Sub
    End If
    mLoading = True
    cboVenue.ListIndex = -1
    mLoading = False
    NewMap
    mDirty = True
    lblStatus.Caption = "Enter a venue name and add tables, then click Save."
    On Error Resume Next
    txtName.SetFocus
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub cmdDelete_Click()
    If mMap.Id = "new" Then
        mDirty = False
        If Len(mActiveId) > 0 Then LoadVenue mActiveId
        Exit Sub
    End If
    If mMap.Id = mActiveId Then
        MsgBox "This seat map is in use and cannot be deleted. Use another venue first.", vbInformation, "Seat Map"
        Exit Sub
    End If
    If MsgBox("Delete venue " & mMap.Name & "?", vbQuestion + vbOKCancel, "Seat Map") <> vbOK Then Exit Sub
    frmMain.SendLine "MAPDEL|" & mMap.Id
    mDirty = False
    mMap.Id = ""
    mWantId = mActiveId
    mLoadAfterList = True
    lblStatus.Caption = "Deleting..."
End Sub

' Checks the map makes sense before it goes to the server
Private Function MapIsOK() As Boolean
    Dim t As Long, k As Long
    If Len(Trim$(mMap.Name)) = 0 Then
        lblStatus.Caption = "Enter a venue name."
        Beep
        Exit Function
    End If
    For t = 0 To mMap.Count - 1
        If Len(mMap.Tables(t).Prefix) = 0 Then
            FillTables t
            lblStatus.Caption = "Enter a seat letter for " & mMap.Tables(t).Name & " (for example, " & FreeLetter() & ")."
            Beep
            Exit Function
        End If
        For k = 0 To t - 1
            If mMap.Tables(k).Prefix = mMap.Tables(t).Prefix Then
                FillTables t
                lblStatus.Caption = mMap.Tables(k).Name & " already uses the seat letter " & mMap.Tables(t).Prefix & "."
                Beep
                Exit Function
            End If
        Next
    Next
    MapIsOK = True
End Function

Private Sub SaveVenue()
    If Not MapIsOK() Then
        mUseAfterSave = False
        Exit Sub
    End If
    lblStatus.Caption = "Saving..."
    frmMain.SendLine "MAPSAVE|" & mMap.Id & "|" & MapSafe(mMap.Name) & "|" & MapSafe(mMap.Stage) & "|" & mMap.Cols & "|" & MapTablesText(mMap)
End Sub

Private Sub UseVenue()
    mUsing = True
    lblStatus.Caption = "Activating " & mMap.Name & "..."
    frmMain.SendLine "MAPUSE|" & mMap.Id
End Sub

Private Sub cmdSave_Click()
    mUseAfterSave = False
    SaveVenue
End Sub

Private Sub cmdUse_Click()
    If mDirty Or mMap.Id = "new" Or Len(mMap.Id) = 0 Then
        mUseAfterSave = True
        SaveVenue
    Else
        UseVenue
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' ---- small helpers

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
