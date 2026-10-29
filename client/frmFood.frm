VERSION 5.00
Begin VB.Form frmFood
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Food Order"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   LinkTopic       =   "frmFood"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4665
      TabIndex        =   14
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame fraHost
      Caption         =   "Host"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   5520
      Begin VB.CommandButton cmdCopy
         Caption         =   "Cop&y"
         Height          =   375
         Left            =   4770
         TabIndex        =   13
         ToolTipText     =   "Copy all the orders to the clipboard"
         Top             =   300
         Width           =   645
      End
      Begin VB.CommandButton cmdClear
         Caption         =   "Clear &All"
         Height          =   375
         Left            =   3915
         TabIndex        =   12
         ToolTipText     =   "Remove all the orders"
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton cmdOpen
         Caption         =   "&Open Orders"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtTitle
         Height          =   315
         Left            =   735
         MaxLength       =   60
         TabIndex        =   10
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label lblOrderName
         Caption         =   "N&ame:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.ListBox lstAll
      Height          =   1620
      Left            =   240
      TabIndex        =   7
      Top             =   2475
      Width           =   5520
   End
   Begin VB.Frame fraMine
      Caption         =   "Your Order"
      Height          =   1155
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5520
      Begin VB.CommandButton cmdCancelOrder
         Caption         =   "Ca&ncel Order"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   690
         Width           =   1455
      End
      Begin VB.CommandButton cmdPlace
         Caption         =   "&Place Order"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   1455
      End
      Begin VB.TextBox txtMine
         Height          =   315
         Left            =   120
         MaxLength       =   80
         TabIndex        =   3
         Top             =   300
         Width           =   5280
      End
   End
   Begin VB.Label lblAll
      Caption         =   "All orders:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2235
      Width           =   5520
   End
   Begin VB.Label lblState
      ForeColor       =   &H00636163&
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   450
      Width           =   5520
   End
   Begin VB.Label lblTitle
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   5520
   End
End
Attribute VB_Name = "frmFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - What's On > Food Order.
Option Explicit

Private mFilling As Boolean     ' Your Order is being filled in, not typed in
Private mTyped As Boolean       ' Your Order has been typed in since it was filled in

Private Sub Form_Load()
    Dim delta As Long
    mFilling = False
    mTyped = False
    lblTitle.FontBold = True
    If frmMain.HostMode() Then
        fraHost.Visible = True
    Else
        ' no Host box: close up the gap
        fraHost.Visible = False
        delta = cmdClose.Top - fraHost.Top
        cmdClose.Top = fraHost.Top
        Me.Height = Me.Height - delta
    End If
    RefreshFood
End Sub

' Called again by the main window whenever an order comes in
Public Sub RefreshFood()
    Dim i As Long, n As Long, isOpen As Boolean, mine As String
    isOpen = (frmMain.FoodState() = "open")
    n = frmMain.FoodCount()
    If Len(frmMain.FoodTitle()) > 0 Then
        lblTitle.Caption = frmMain.FoodTitle()
    Else
        lblTitle.Caption = "Food Order"
    End If
    lblState.ForeColor = &H636163
    If isOpen Then
        lblState.Caption = "Orders are open. Your order can be changed until the host closes the order."
    ElseIf n > 0 Then
        lblState.Caption = "Orders are closed."
    Else
        lblTitle.Caption = "No food order is open."
        lblState.Caption = "The host opens a food order when it is time to order. It is announced in Chat."
    End If
    mine = frmMain.MyFood()
    If Not mTyped Then
        mFilling = True
        txtMine.Text = mine
        mFilling = False
    End If
    txtMine.Enabled = isOpen
    cmdPlace.Enabled = isOpen
    If Len(mine) > 0 Then
        cmdPlace.Caption = "&Change Order"
    Else
        cmdPlace.Caption = "&Place Order"
    End If
    cmdCancelOrder.Enabled = isOpen And Len(mine) > 0
    lstAll.Clear
    For i = 0 To n - 1
        lstAll.AddItem frmMain.FoodNick(i) & " (" & AmPm(frmMain.FoodTime(i)) & "):  " & frmMain.FoodText(i)
    Next
    lblAll.Caption = "All orders (" & n & "):"
    ' the host's box
    If isOpen Then
        cmdOpen.Caption = "C&lose Orders"
    Else
        cmdOpen.Caption = "&Open Orders"
    End If
    If Len(txtTitle.Text) = 0 Then txtTitle.Text = frmMain.FoodTitle()
    cmdClear.Enabled = (n > 0)
    cmdCopy.Enabled = (n > 0)
End Sub

' A message from the server (the order has closed)
Public Sub ShowError(ByVal msg As String)
    lblState.ForeColor = vbRed
    lblState.Caption = msg
    Beep
End Sub

Private Sub txtMine_Change()
    If Not mFilling Then mTyped = True
End Sub

Private Sub cmdPlace_Click()
    Dim s As String
    s = Trim$(Replace(txtMine.Text, "|", "/"))
    If Len(s) = 0 Then
        ShowError "Enter your order (for example, large pepperoni and a Coke)."
        txtMine.SetFocus
        Exit Sub
    End If
    mTyped = False
    frmMain.SendLine "FOOD|" & s
End Sub

Private Sub cmdCancelOrder_Click()
    mTyped = False
    frmMain.SendLine "FOOD|"
End Sub

' ---- the host

Private Sub cmdOpen_Click()
    Dim s As String
    If frmMain.FoodState() = "open" Then
        frmMain.SendLine "FOODCLOSE"
    Else
        s = Trim$(Replace(txtTitle.Text, "|", "/"))
        If Len(s) = 0 Then s = "Food order"
        frmMain.SendLine "FOODOPEN|" & s
    End If
End Sub

Private Sub cmdClear_Click()
    If MsgBox("Remove all " & frmMain.FoodCount() & " orders?", vbQuestion + vbOKCancel + vbDefaultButton2, "Food Order") = vbOK Then
        frmMain.SendLine "FOODCLEAR"
    End If
End Sub

Private Sub cmdCopy_Click()
    Dim s As String, i As Long
    s = frmMain.FoodTitle() & vbCrLf & vbCrLf
    For i = 0 To frmMain.FoodCount() - 1
        s = s & frmMain.FoodNick(i) & ": " & frmMain.FoodText(i) & vbCrLf
    Next
    s = s & vbCrLf & frmMain.FoodCount() & " orders" & vbCrLf
    Clipboard.Clear
    Clipboard.SetText s
    lblState.ForeColor = &H636163
    lblState.Caption = "The order list has been copied to the clipboard."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
