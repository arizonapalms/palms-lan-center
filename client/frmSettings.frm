VERSION 5.00
Begin VB.Form frmSettings
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "frmSettings"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      Top             =   3540
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   29
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Frame fraChat
      Caption         =   "Chat"
      Height          =   615
      Left            =   240
      TabIndex        =   26
      Top             =   2760
      Width           =   4920
      Begin VB.CheckBox chkJoins
         Caption         =   "Show &sign-in, sign-out and away messages"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   255
         Width           =   4440
      End
   End
   Begin VB.Frame fraGames
      Caption         =   "Game Library"
      Height          =   1095
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   4920
      Begin VB.CommandButton cmdGames
         Caption         =   "&Browse..."
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   255
         Width           =   1000
      End
      Begin VB.TextBox txtGames
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   255
         Width           =   2175
      End
      Begin VB.Label lblGamesNote
         Caption         =   "Applies to new installations. Installed games are not moved."
         ForeColor       =   &H00636163&
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   660
         Width           =   4440
      End
      Begin VB.Label lblGames
         Caption         =   "&Install games to:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame fraSound
      Caption         =   "Sounds"
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   4920
      Begin VB.PictureBox picVol
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   1080
         ScaleMode       =   3  'Pixel
         TabIndex        =   18
         Top             =   660
         Width           =   2880
      End
      Begin VB.CheckBox chkSounds
         Caption         =   "&Play sounds"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Width           =   4440
      End
      Begin VB.Label lblVolPct
         Caption         =   "80%"
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   735
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblVol
         Caption         =   "&Volume:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   735
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - Tools > Options: sounds, chat messages and the games folder.
Option Explicit

Private mVol As Long
Private mDragging As Boolean
Private mFocused As Boolean

Private Sub Form_Load()
    chkSounds.Value = CheckOf(gSoundOn)
    chkJoins.Value = CheckOf(gShowJoins)
    mVol = gVolume
    DrawVol
    txtGames.Text = GamesRoot()
End Sub

Private Function CheckOf(ByVal b As Boolean) As Integer
    If b Then
        CheckOf = vbChecked
    Else
        CheckOf = vbUnchecked
    End If
End Function

' ---- the volume slider, drawn like a Windows 98 one

Private Sub DrawVol()
    Dim w As Long, h As Long, track As Long, k As Long, x As Long, tx As Long
    picVol.Cls
    w = picVol.ScaleWidth
    h = picVol.ScaleHeight
    track = w - 12 - 11
    Sunken2 picVol, 6, 7, w - 12, 4
    For k = 0 To 10
        x = 6 + 5 + (track * k) \ 10
        Box picVol, x, h - 6, 1, 4, vb3DShadow
    Next
    tx = 6 + (track * mVol) \ 100
    Box picVol, tx, 0, 11, h - 8, vbButtonFace
    Edge picVol, tx, 0, 11, h - 8, vb3DHighlight, vb3DDKShadow
    Edge picVol, tx + 1, 1, 9, h - 10, vb3DLight, vb3DShadow
    If mFocused Then
        picVol.DrawStyle = vbDot
        picVol.Line (tx - 2, 0)-(tx + 12, h - 8), SysCol(vbButtonText), B
        picVol.DrawStyle = vbSolid
    End If
    lblVolPct.Caption = mVol & "%"
End Sub

Private Sub SetVol(ByVal v As Long)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    If v <> mVol Then
        mVol = v
        DrawVol
    End If
End Sub

Private Sub VolFromX(ByVal px As Single)
    Dim track As Long, v As Long
    track = picVol.ScaleWidth - 12 - 11
    If track < 1 Then Exit Sub
    v = CLng((px - 11) * 100 / track)
    v = ((v + 2) \ 5) * 5          ' in steps of 5%
    SetVol v
End Sub

Private Sub picVol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    mDragging = True
    VolFromX X
End Sub

Private Sub picVol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mDragging Then VolFromX X
End Sub

Private Sub picVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mDragging Then
        mDragging = False
        PreviewSound
    End If
End Sub

Private Sub picVol_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft, vbKeyDown
        SetVol mVol - 5
    Case vbKeyRight, vbKeyUp
        SetVol mVol + 5
    Case vbKeyHome
        SetVol 0
    Case vbKeyEnd
        SetVol 100
    End Select
End Sub

Private Sub picVol_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft, vbKeyDown, vbKeyRight, vbKeyUp, vbKeyHome, vbKeyEnd
        PreviewSound
    End Select
End Sub

Private Sub picVol_GotFocus()
    mFocused = True
    DrawVol
End Sub

Private Sub picVol_LostFocus()
    mFocused = False
    DrawVol
End Sub

' Plays a ding at the slider's volume, before you've pressed OK
Private Sub PreviewSound()
    Dim oldVol As Long, oldOn As Boolean
    oldVol = gVolume
    oldOn = gSoundOn
    gVolume = mVol
    gSoundOn = (chkSounds.Value = vbChecked)
    PlayChime "ding"
    gVolume = oldVol
    gSoundOn = oldOn
End Sub

' ---- buttons

Private Sub cmdGames_Click()
    Dim s As String
    s = PickFolder(Me.hWnd, "Select the folder to install games to:")
    If Len(s) > 0 Then txtGames.Text = s
End Sub

Private Sub cmdOK_Click()
    Dim games As String
    ' the games folder: a folder on this PC, like C:\LANPartyGames or D:\Games
    games = Trim$(txtGames.Text)
    If Len(games) > 3 And Right$(games, 1) = "\" Then games = Left$(games, Len(games) - 1)
    If Len(games) = 2 And Right$(games, 1) = ":" Then games = games & "\"
    If Not (games Like "[A-Za-z]:\*") Or InStr(games, "..") > 0 Then
        MsgBox "Select a folder on a local drive (for example, C:\LANPartyGames or D:\Games).", vbExclamation, "Options"
        On Error Resume Next
        txtGames.SetFocus
        Exit Sub
    End If
    SettingSet "games_folder", games
    gSoundOn = (chkSounds.Value = vbChecked)
    gShowJoins = (chkJoins.Value = vbChecked)
    gVolume = mVol
    SaveSettings
    frmMain.SettingsChanged
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
