VERSION 5.00
Begin VB.Form frmNotify
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notifications"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "frmNotify"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   6630
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   6630
      Width           =   1095
   End
   Begin VB.CommandButton cmdTest
      Caption         =   "&Test"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Show a test notification in the selected style"
      Top             =   6630
      Width           =   1095
   End
   Begin VB.Frame fraStyle
      Caption         =   "Notification Style"
      Height          =   3495
      Left            =   240
      TabIndex        =   8
      Top             =   2970
      Width           =   6120
      Begin VB.PictureBox picClassic
         AutoRedraw      =   -1  'True
         Height          =   1530
         Left            =   3480
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Classic message box"
         Top             =   600
         Width           =   2280
      End
      Begin VB.PictureBox picMsn
         AutoRedraw      =   -1  'True
         Height          =   1530
         Left            =   480
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Messenger style"
         Top             =   600
         Width           =   2280
      End
      Begin VB.CheckBox chkInterrupt
         Caption         =   "&Show on top of full-screen games"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   5640
      End
      Begin VB.OptionButton optClassic
         Caption         =   "C&lassic message box"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   300
         Width           =   2640
      End
      Begin VB.OptionButton optMessenger
         Caption         =   "&Messenger style"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   2640
      End
      Begin VB.Label lblInterrupt
         Caption         =   "Not recommended. Some full-screen games are minimised when a notification appears on top."
         ForeColor       =   &H00636163&
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   3030
         Width           =   5400
      End
      Begin VB.Label lblClassic
         Caption         =   "Stays open in the centre of the screen until OK is clicked."
         ForeColor       =   &H00636163&
         Height          =   435
         Left            =   3240
         TabIndex        =   15
         Top             =   2220
         Width           =   2640
      End
      Begin VB.Label lblMsn
         Caption         =   "Appears above the taskbar and closes after a few seconds."
         ForeColor       =   &H00636163&
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   2220
         Width           =   2760
      End
   End
   Begin VB.Frame fraPop
      Caption         =   "Show Notifications For"
      Height          =   2745
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6120
      Begin VB.CheckBox chkCups
         Caption         =   "T&ournaments: sign-ups, my matches and the winner"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   5640
      End
      Begin VB.CheckBox chkNext
         Caption         =   "&Upcoming events (15, 10 and 5 minutes before)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1950
         Width           =   5640
      End
      Begin VB.CheckBox chkName
         Caption         =   "My &name in Chat"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1620
         Width           =   5640
      End
      Begin VB.CheckBox chkResult
         Caption         =   "Vote &results"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1290
         Width           =   5640
      End
      Begin VB.CheckBox chkSoon
         Caption         =   "Vote &closing soon"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   5640
      End
      Begin VB.CheckBox chkVote
         Caption         =   "&Vote started"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   630
         Width           =   5640
      End
      Begin VB.CheckBox chkPA
         Caption         =   "&Announcements from the host"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   5640
      End
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Tools > Notifications: which pop-ups appear, and their style.
Option Explicit

Private Const RES_MSN As Long = 105
Private Const RES_CLASSIC As Long = 106

Private Sub Form_Load()
    chkPA.Value = CheckOf(gNotifyPA)
    chkVote.Value = CheckOf(gNotifyVote)
    chkSoon.Value = CheckOf(gNotifySoon)
    chkResult.Value = CheckOf(gNotifyResult)
    chkName.Value = CheckOf(gNotifyName)
    chkNext.Value = CheckOf(gNotifyNext)
    chkCups.Value = CheckOf(gNotifyCups)
    chkInterrupt.Value = CheckOf(gInterrupt)
    If gClassicPopups Then
        optClassic.Value = True
    Else
        optMessenger.Value = True
    End If
    On Error Resume Next
    picMsn.Picture = LoadResPicture(RES_MSN, vbResBitmap)
    picClassic.Picture = LoadResPicture(RES_CLASSIC, vbResBitmap)
End Sub

Private Function CheckOf(ByVal b As Boolean) As Integer
    If b Then
        CheckOf = vbChecked
    Else
        CheckOf = vbUnchecked
    End If
End Function

' Clicking a picture picks that style
Private Sub picMsn_Click()
    optMessenger.Value = True
End Sub

Private Sub picClassic_Click()
    optClassic.Value = True
End Sub

Private Sub cmdTest_Click()
    Dim oldClassic As Boolean, oldInterrupt As Boolean
    oldClassic = gClassicPopups
    oldInterrupt = gInterrupt
    gClassicPopups = optClassic.Value
    gInterrupt = (chkInterrupt.Value = vbChecked)
    PlayChime "ding"
    frmMain.ShowToast "test", IC_STAR, "Test Notification", "Notifications appear like this. Click a notification to open Palms LAN Center.", 6
    gClassicPopups = oldClassic
    gInterrupt = oldInterrupt
End Sub

Private Sub cmdOK_Click()
    gNotifyPA = (chkPA.Value = vbChecked)
    gNotifyVote = (chkVote.Value = vbChecked)
    gNotifySoon = (chkSoon.Value = vbChecked)
    gNotifyResult = (chkResult.Value = vbChecked)
    gNotifyName = (chkName.Value = vbChecked)
    gNotifyNext = (chkNext.Value = vbChecked)
    gNotifyCups = (chkCups.Value = vbChecked)
    gInterrupt = (chkInterrupt.Value = vbChecked)
    gClassicPopups = optClassic.Value
    SaveSettings
    frmMain.SettingsChanged
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
