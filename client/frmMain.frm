VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain
   Caption         =   "Palms LAN Center"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleMode       =   3  'Pixel
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClock
      Interval        =   1000
      Left            =   200
      Top             =   440
   End
   Begin VB.Timer tmrNote
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8
      Top             =   440
   End
   Begin VB.Timer tmrHover
      Enabled         =   0   'False
      Interval        =   100
      Left            =   40
      Top             =   440
   End
   Begin VB.Timer tmrReconnect
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   72
      Top             =   440
   End
   Begin VB.Timer tmrPing
      Interval        =   30000
      Left            =   104
      Top             =   440
   End
   Begin VB.Timer tmrLazy
      Enabled         =   0   'False
      Interval        =   50
      Left            =   232
      Top             =   440
   End
   Begin VB.Timer tmrLib
      Enabled         =   0   'False
      Interval        =   500
      Left            =   264
      Top             =   440
   End
   Begin MSWinsockLib.Winsock sckFind
      Left            =   136
      Top             =   440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckTCP
      Left            =   168
      Top             =   440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picTools
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      ScaleMode       =   3  'Pixel
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   770
   End
   Begin VB.PictureBox picStatus
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   22
      Left            =   0
      ScaleMode       =   3  'Pixel
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   488
      Width           =   770
   End
   Begin VB.PictureBox picSide
      AutoRedraw      =   -1  'True
      BackColor       =   &H006B6131&
      BorderStyle     =   0  'None
      Height          =   454
      Left            =   2
      ScaleMode       =   3  'Pixel
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   32
      Width           =   88
   End
   Begin VB.PictureBox picPane
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   454
      Left            =   94
      ScaleMode       =   3  'Pixel
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   32
      Width           =   674
      Begin VB.PictureBox picHeader
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   53
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2
         Width           =   670
      End
      Begin VB.PictureBox picSoonPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
      End
      Begin VB.PictureBox picWhatsPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.PictureBox picVote
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   380
            ScaleMode       =   3  'Pixel
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   28
            Width           =   270
         End
         Begin VB.PictureBox picSched
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   16
            ScaleMode       =   3  'Pixel
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   28
            Width           =   340
            Begin VB.VScrollBar vsbSched
               Height          =   296
               Left            =   320
               Max             =   0
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
      End
      Begin VB.PictureBox picWhoPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.CommandButton cmdWhoChat
            Caption         =   "&Chat"
            Height          =   25
            Left            =   420
            TabIndex        =   55
            Top             =   300
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton cmdWhoSeat
            Caption         =   "Change &Seat..."
            Height          =   25
            Left            =   420
            TabIndex        =   54
            Top             =   300
            Visible         =   0   'False
            Width           =   130
         End
         Begin VB.PictureBox picMap
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E7F7FF&
            Height          =   300
            Left            =   16
            ScaleMode       =   3  'Pixel
            TabIndex        =   56
            Top             =   34
            Visible         =   0   'False
            Width           =   380
            Begin VB.VScrollBar vsbMap
               Height          =   296
               Left            =   360
               Max             =   0
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
         Begin VB.PictureBox picCard
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   260
            Left            =   420
            ScaleMode       =   3  'Pixel
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   28
            Width           =   230
         End
         Begin VB.PictureBox picPeople
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   16
            ScaleMode       =   3  'Pixel
            TabIndex        =   51
            Top             =   28
            Width           =   380
            Begin VB.VScrollBar vsbPeople
               Height          =   296
               Left            =   360
               Max             =   0
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
      End
      Begin VB.PictureBox picLibPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.CommandButton cmdLibDel
            Caption         =   "&Uninstall"
            Height          =   25
            Left            =   530
            TabIndex        =   63
            Top             =   300
            Visible         =   0   'False
            Width           =   100
         End
         Begin VB.CommandButton cmdLibAlt
            Caption         =   "&Update"
            Height          =   25
            Left            =   420
            TabIndex        =   62
            Top             =   300
            Visible         =   0   'False
            Width           =   100
         End
         Begin VB.CommandButton cmdLibGo
            Caption         =   "&Install"
            Height          =   25
            Left            =   420
            TabIndex        =   61
            Top             =   330
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.PictureBox picLibCard
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   260
            Left            =   420
            ScaleMode       =   3  'Pixel
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   28
            Width           =   230
         End
         Begin VB.PictureBox picLibList
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   16
            ScaleMode       =   3  'Pixel
            TabIndex        =   59
            ToolTipText     =   "Double-click a game to install or play it"
            Top             =   28
            Width           =   380
            Begin VB.VScrollBar vsbLib
               Height          =   296
               Left            =   360
               Max             =   0
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
      End
      Begin VB.PictureBox picCupsPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.CommandButton cmdCupGo
            Caption         =   "&Sign Up"
            Height          =   25
            Left            =   520
            TabIndex        =   69
            Top             =   360
            Visible         =   0   'False
            Width           =   140
         End
         Begin VB.PictureBox picCup
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   16
            ScaleMode       =   3  'Pixel
            TabIndex        =   66
            Top             =   33
            Width           =   600
            Begin VB.HScrollBar hsbCup
               Height          =   16
               Left            =   0
               Max             =   0
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   280
               Visible         =   0   'False
               Width           =   580
            End
            Begin VB.VScrollBar vsbCup
               Height          =   280
               Left            =   580
               Max             =   0
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   16
            End
         End
      End
      Begin VB.PictureBox picPAPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.CommandButton cmdPASend
            Caption         =   "&Announce"
            Height          =   25
            Left            =   126
            TabIndex        =   26
            Top             =   178
            Width           =   120
         End
         Begin VB.CommandButton cmdPATest
            Caption         =   "&Test Sound"
            Height          =   25
            Left            =   10
            TabIndex        =   25
            Top             =   178
            Width           =   110
         End
         Begin VB.OptionButton optPA
            Caption         =   "None"
            Height          =   18
            Index           =   3
            Left            =   180
            TabIndex        =   24
            Top             =   141
            Width           =   140
         End
         Begin VB.OptionButton optPA
            Caption         =   "Air horn"
            Height          =   18
            Index           =   2
            Left            =   24
            TabIndex        =   23
            Top             =   141
            Width           =   140
         End
         Begin VB.OptionButton optPA
            Caption         =   "Triple chime"
            Height          =   18
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   120
            Width           =   140
         End
         Begin VB.OptionButton optPA
            Caption         =   "Ding-dong"
            Height          =   18
            Index           =   0
            Left            =   24
            TabIndex        =   21
            Top             =   120
            Value           =   -1  'True
            Width           =   140
         End
         Begin VB.TextBox txtPA
            Height          =   58
            Left            =   10
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   28
            Width           =   400
         End
         Begin VB.Label lblPA
            BackStyle       =   0  'Transparent
            Caption         =   "&Message:"
            Height          =   16
            Left            =   10
            TabIndex        =   19
            Top             =   10
            Width           =   400
         End
      End
      Begin VB.PictureBox picChatPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   55
         Visible         =   0   'False
         Width           =   670
         Begin VB.PictureBox picNext
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E7FFFF&
            ForeColor       =   &H00000000&
            Height          =   74
            Left            =   482
            ScaleMode       =   3  'Pixel
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   313
            Width           =   178
         End
         Begin VB.PictureBox picWho
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   269
            Left            =   488
            ScaleMode       =   3  'Pixel
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Double-click a name to view details"
            Top             =   28
            Width           =   166
            Begin VB.VScrollBar vsbWho
               Height          =   265
               Left            =   146
               Max             =   0
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
         Begin VB.CommandButton cmdSend
            Caption         =   "&Send"
            Height          =   23
            Left            =   382
            TabIndex        =   12
            Top             =   364
            Width           =   80
         End
         Begin VB.PictureBox picSmiley
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   23
            Left            =   355
            ScaleMode       =   3  'Pixel
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Insert smiley"
            Top             =   364
            Width           =   23
         End
         Begin VB.TextBox txtSay
            Height          =   23
            Left            =   10
            MaxLength       =   200
            TabIndex        =   10
            Top             =   364
            Width           =   341
         End
         Begin VB.PictureBox picLog
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   348
            Left            =   10
            ScaleMode       =   3  'Pixel
            TabIndex        =   39
            Top             =   10
            Width           =   452
            Begin VB.VScrollBar vsbLog
               Height          =   344
               Left            =   432
               Max             =   0
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   0
               Width           =   16
            End
         End
      End
      Begin VB.PictureBox picSignInPage
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   397
         Left            =   2
         ScaleMode       =   3  'Pixel
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   55
         Width           =   670
         Begin VB.CommandButton cmdQuit
            Caption         =   "E&xit"
            Height          =   25
            Left            =   419
            TabIndex        =   7
            Top             =   180
            Width           =   90
         End
         Begin VB.CommandButton cmdSignIn
            Caption         =   "&Sign In"
            Default         =   -1  'True
            Height          =   25
            Left            =   321
            TabIndex        =   6
            Top             =   180
            Width           =   90
         End
         Begin VB.TextBox txtServer
            Height          =   21
            Left            =   265
            MaxLength       =   64
            TabIndex        =   5
            Top             =   139
            Width           =   150
         End
         Begin VB.TextBox txtSeat
            Height          =   21
            Left            =   265
            MaxLength       =   4
            TabIndex        =   3
            Top             =   110
            Width           =   50
         End
         Begin VB.TextBox txtNick
            Height          =   21
            Left            =   265
            MaxLength       =   12
            TabIndex        =   1
            Top             =   81
            Width           =   150
         End
         Begin VB.Label lblServer
            BackStyle       =   0  'Transparent
            Caption         =   "Se&rver:"
            Height          =   16
            Left            =   151
            TabIndex        =   4
            Top             =   142
            Width           =   110
         End
         Begin VB.Label lblSeat
            BackStyle       =   0  'Transparent
            Caption         =   "S&eat (optional):"
            Height          =   16
            Left            =   151
            TabIndex        =   2
            Top             =   113
            Width           =   110
         End
         Begin VB.Label lblNick
            BackStyle       =   0  'Transparent
            Caption         =   "&Nickname:"
            Height          =   16
            Left            =   151
            TabIndex        =   0
            Top             =   84
            Width           =   110
         End
         Begin VB.Label lblFind
            BackStyle       =   0  'Transparent
            Caption         =   "Searching for the Palms Server..."
            Height          =   30
            Left            =   151
            TabIndex        =   45
            Top             =   42
            UseMnemonic     =   0   'False
            Width           =   368
         End
      End
   End
   Begin VB.Menu mnuFile
      Caption         =   "&File"
      Begin VB.Menu mnuFileCopy
         Caption         =   "&Copy Chat"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSave
         Caption         =   "Sa&ve Chat"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrint
         Caption         =   "&Print Schedule..."
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep2
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView
      Caption         =   "&View"
      Begin VB.Menu mnuGo
         Caption         =   "&Chat"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuGo
         Caption         =   "&What's On"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuGo
         Caption         =   "&Play Now"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuGo
         Caption         =   "Game &Library"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuGo
         Caption         =   "W&ho's Here"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuGo
         Caption         =   "&Tournaments"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuGo
         Caption         =   "&Jukebox"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuGo
         Caption         =   "&Announce"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewProj
         Caption         =   "P&rojector"
         Enabled         =   0   'False
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuTools
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSounds
         Caption         =   "&Sounds"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolsNotify
         Caption         =   "&Notifications..."
      End
      Begin VB.Menu mnuToolsOptions
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuToolsSep0
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCart
         Caption         =   "&Make Cartridge..."
      End
      Begin VB.Menu mnuToolsSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsHost
         Caption         =   "&Host Mode..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAdmin
      Caption         =   "H&ost"
      Visible         =   0   'False
      Begin VB.Menu mnuAdminPA
         Caption         =   "Send &Announcement..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminSched
         Caption         =   "Add &Event..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminTpl
         Caption         =   "Schedule &Templates..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminMap
         Caption         =   "S&eat Map..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminVote
         Caption         =   "Start &Vote..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminVoteStop
         Caption         =   "&Close Vote"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminSep2
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminCup
         Caption         =   "New To&urnament..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdminSep3
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminProj
         Caption         =   "&Projector Settings..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpPerf
         Caption         =   "&Diagnostics..."
      End
      Begin VB.Menu mnuHelpSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout
         Caption         =   "&About Palms LAN Center"
      End
   End
   Begin VB.Menu mnuMe
      Caption         =   "Me"
      Visible         =   0   'False
      Begin VB.Menu mnuMeHere
         Caption         =   "&Online"
      End
      Begin VB.Menu mnuMeBreak
         Caption         =   "&Away"
      End
      Begin VB.Menu mnuMeBusy
         Caption         =   "&Busy"
      End
      Begin VB.Menu mnuMeSep0
         Caption         =   "-"
      End
      Begin VB.Menu mnuMeSeat
         Caption         =   "Change &Seat..."
      End
      Begin VB.Menu mnuMeSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuMeSignOut
         Caption         =   "&Sign Out"
      End
   End
   Begin VB.Menu mnuStat
      Caption         =   "Status"
      Visible         =   0   'False
      Begin VB.Menu mnuStatOnline
         Caption         =   "&Online"
      End
      Begin VB.Menu mnuStatAway
         Caption         =   "&Away"
      End
      Begin VB.Menu mnuStatBusy
         Caption         =   "&Busy"
      End
   End
   Begin VB.Menu mnuSmileys
      Caption         =   "Smileys"
      Visible         =   0   'False
      Begin VB.Menu mnuSmile
         Caption         =   ":)"
         Index           =   0
      End
      Begin VB.Menu mnuSmile
         Caption         =   ":D"
         Index           =   1
      End
      Begin VB.Menu mnuSmile
         Caption         =   ";)"
         Index           =   2
      End
      Begin VB.Menu mnuSmile
         Caption         =   ":P"
         Index           =   3
      End
      Begin VB.Menu mnuSmile
         Caption         =   ":O"
         Index           =   4
      End
      Begin VB.Menu mnuSmile
         Caption         =   "<3"
         Index           =   5
      End
      Begin VB.Menu mnuSmile
         Caption         =   "gg"
         Index           =   6
      End
      Begin VB.Menu mnuSmile
         Caption         =   "o/"
         Index           =   7
      End
   End
   Begin VB.Menu mnuEvent
      Caption         =   "Event"
      Visible         =   0   'False
      Begin VB.Menu mnuEvIn
         Caption         =   "&Count Me In"
      End
      Begin VB.Menu mnuEvRemind
         Caption         =   "&Remind Me"
      End
      Begin VB.Menu mnuEvInfo
         Caption         =   "&Details..."
      End
      Begin VB.Menu mnuEvChat
         Caption         =   "Copy to C&hat"
      End
      Begin VB.Menu mnuEvSep
         Caption         =   "-"
      End
      Begin VB.Menu mnuEvDel
         Caption         =   "Re&move Event"
      End
   End
   Begin VB.Menu mnuMore
      Caption         =   "More"
      Visible         =   0   'False
      Begin VB.Menu mnuMoreItem
         Caption         =   "(none)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - The main window: sidebar, toolbar, pages and status bar, all drawn by hand.
Option Explicit

Private Const DEFAULT_PORT As Long = 4098
Private Const BEACON_PORT As Long = 4099

' Pages. The sidebar shows them in this order.
Private Const PG_SIGNIN As Integer = -1
Private Const PG_CHAT As Integer = 0
Private Const PG_WHATSON As Integer = 1
Private Const PG_PLAY As Integer = 2
Private Const PG_LIBRARY As Integer = 3
Private Const PG_WHO As Integer = 4
Private Const PG_CUPS As Integer = 5
Private Const PG_JUKEBOX As Integer = 6
Private Const PG_PA As Integer = 7
Private Const PG_COUNT As Integer = 8

' Sizes, in pixels
Private Const ROW_H As Long = 17
Private Const TIME_X As Long = 6
Private Const ICON_X As Long = 40
Private Const TEXT_X As Long = 61
Private Const NAV_TOP As Long = 6
Private Const NAV_H As Long = 58          ' sidebar button height
Private Const NAV_GAP As Long = 2
Private Const NAV_ARROW_H As Long = 16    ' the sidebar's scroll arrows
Private Const MAX_LINES As Long = 300
Private Const MAX_TOOLS As Integer = 24
Private Const ME_TOOL As Integer = 99      ' the "Signed in as" button on the toolbar
Private Const CHEV_TOOL As Integer = 98    ' the >> button: buttons that don't fit, as a menu
Private Const CHEV_W As Long = 22
Private Const REMIND_MINS As Long = 5      ' Remind Me pops up this many minutes before the event
Private Const SCHED_ROW As Long = 24
Private Const VOTE_TOP As Long = 48
Private Const VOTE_ROW As Long = 38
Private Const PPL_ROW As Long = 18         ' Who's Here list rows
Private Const HDR_H As Long = 19           ' Who's Here column headings
Private Const COL_SEAT_W As Long = 44
Private Const COL_STATUS_W As Long = 92
Private Const COL_PING_W As Long = 70
Private Const COL_YOU_W As Long = 62
Private Const TAB_BOTTOM As Long = 27     ' where the Who's Here tabs meet their panel
Private Const SLOW_LINK_MB As Long = 100  ' a network card slower than this needs a hand
Private Const SLOW_PING_MS As Long = 50   ' so does a PC this slow to answer the server
Private Const MAX_PEERS As Long = 64

' Parts of the window waiting to be redrawn. Anything that changes asks for a redraw;
' tmrLazy then draws everything that was asked for, once, and only what's on screen.
Private Const DR_TOOLS As Long = 1
Private Const DR_SIDE As Long = 2
Private Const DR_STATUS As Long = 4
Private Const DR_HEADER As Long = 8
Private Const DR_LOG As Long = 16          ' Chat: the conversation
Private Const DR_WHO As Long = 32          ' Chat: the who's here box
Private Const DR_NEXT As Long = 64         ' Chat: the coming up box
Private Const DR_WHATS As Long = 128       ' What's On: the whole page
Private Const DR_SCHED As Long = 256
Private Const DR_VOTE As Long = 512
Private Const DR_WHOPAGE As Long = 1024    ' Who's Here: the whole page
Private Const DR_WHOBODY As Long = 2048    ' Who's Here: the list or the map
Private Const DR_CARD As Long = 4096       ' Who's Here: the card
Private Const DR_LIBPAGE As Long = 8192    ' Game Library: the whole page
Private Const DR_LIBLIST As Long = 16384   ' Game Library: the list
Private Const DR_LIBCARD As Long = 32768   ' Game Library: the card
Private Const DR_CUPPAGE As Long = 65536   ' Tournaments: the whole page
Private Const DR_CUP As Long = 131072      ' Tournaments: the bracket

' The Game Library list
Private Const LIB_ROW As Long = 20
Private Const LIB_PLAYERS_W As Long = 54
Private Const LIB_SIZE_W As Long = 62
Private Const LIB_STATUS_W As Long = 128

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

' ---- talking to the server
Private mBuffer As String
Private mServerHost As String
Private mServerPort As Long
Private mServerName As String
Private mPcId As String
Private mSignedIn As Boolean
Private mMyNick As String
Private mMySeat As String
Private mWantNick As String
Private mWantSeat As String
Private mPendingLogin As Boolean
Private mAutoRelogin As Boolean
Private mIsAdmin As Boolean
Private mAway As Boolean
Private mBusy As Boolean
Private mSigningOut As Boolean
Private mClosing As Boolean
Private mPingSent As Long
Private mPingMs As Long
Private mServerVer As String


' ---- the look
Private mPage As Integer
Private mNavHover As Integer
Private mNavDown As Integer
Private mNavFirst As Integer          ' the first sidebar button shown, when they don't all fit
Private mNavArrowDown As Integer      ' the sidebar arrow being pressed: -1 up, 1 down, 0 none
Private mToolHover As Integer
Private mToolDown As Integer
Private mUnread As Long
Private mConnText As String
Private mConnIcon As Integer
Private mNoteText As String
Private mIgnoreScroll As Boolean

' ---- toolbar buttons for the page on show
Private mToolCount As Integer
Private mToolKey(0 To 23) As String
Private mToolIcon(0 To 23) As Integer
Private mToolText(0 To 23) As String
Private mToolOn(0 To 23) As Boolean
Private mToolL(0 To 23) As Long
Private mToolR(0 To 23) As Long
Private mToolHide(0 To 23) As Boolean     ' didn't fit: it's on the >> menu instead
Private mChevL As Long                    ' the >> button, when there is one
Private mChevR As Long
Private mMoreMap(0 To 23) As Integer      ' >> menu item -> toolbar button

' ---- chat lines
Private mLineCount As Long
Private mLnTime() As String
Private mLnIcon() As Integer
Private mLnNick() As String
Private mLnNickCol() As Long
Private mLnNickW() As Long
Private mLnText() As String
Private mLnCol() As Long
Private mLnBg() As Long
Private mLnWrap() As String
Private mLnRows() As Long

' ---- the "Signed in as" button
Private mMeL As Long
Private mMeR As Long
Private mMeOpen As Boolean
Private mWelcome As String
Private mWelcomeNote As String        ' said after the welcome, like "seat A01 was taken"

' ---- What's On: today's schedule
Private mSchedBuild As Collection
Private mSchedCount As Long
Private mSchedId() As String
Private mSchedTime() As String
Private mSchedIcon() As Integer
Private mSchedName() As String
Private mSchedSel As Long
Private mSchedInfo As Collection      ' SCHEDINFO lines waiting for SCHEDEND
Private mSchedIn() As String          ' who's in: nicknames separated by commas
Private mSchedInN() As Long
Private mSchedNote() As String        ' the host's details for the event, ~ between lines
Private mSchedGame() As String        ' the game it's for (cartridge/game), for the projector
Private mSchedServer() As String      ' the game server's address, like 10.0.200.80:27960
Private mSchedConnect() As String     ' how to connect to it
Private mPendingGame As String        ' game|server|how to connect for the event just added
Private mProjSecs As Long             ' the projector: seconds per slide, the slides, the web page's port
Private mProjSlides As String
Private mProjPort As String
Private mPendingNote As String        ' details for the event just added, sent once the server says its ID
Private mRemind As String             ' Remind Me: id@time;id@time (kept in the registry)
Private mRemindDone As String         ' reminders already popped up
Private mLastMinute As String

' ---- What's On: the next-game vote
Private mVoteState As String          ' none, open or closed
Private mVoteDeadline As Long         ' GetTickCount when it closes
Private mVoteMine As Long             ' which choice I picked, 0 = none yet
Private mVoteQ As String
Private mVoteCount As Long
Private mVoteName(0 To 5) As String
Private mVoteVotes(0 To 5) As Long
Private mVoteTotal As Long
Private mVoteWinner As Long
Private mPrevVoteState As String      ' the state before the latest update, to spot changes
Private mVoteKey As String            ' question and choices of the vote we've seen
Private mSoonAt As Long               ' seconds left when we warn "about to close"
Private mSoonDone As Boolean
Private mLastNextKey As String        ' the last "up next in N minutes" we popped up

' ---- What's On: games suggested for the next vote, and the food order
Private mSuggCount As Long
Private mSuggName(0 To 15) As String
Private mSuggBy(0 To 15) As String    ' nicknames separated by commas
Private mSuggN(0 To 15) As Long
Private mSuggFresh As Boolean         ' the next SUGG line starts a new list
Private mFoodState As String          ' open or closed
Private mFoodTitle As String
Private mFoodCount As Long
Private mFoodNick(0 To 99) As String
Private mFoodTime(0 To 99) As String
Private mFoodText(0 To 99) As String
Private mFoodHeard As Boolean         ' a FOODINFO has come in since signing in

' ---- who's here
Private mWhoBuild As Collection
Private mWhoCount As Long
Private mWhoNick() As String
Private mWhoSeat() As String
Private mWhoAway() As Boolean
Private mWhoBusy() As Boolean
Private mWhoAdmin() As Boolean
Private mWhoSince() As String         ' when they arrived, like 10:02
Private mWhoIp() As String            ' their IP address
Private mWhoMs() As String            ' their ping from the server, "" if not known
Private mWhoLink() As String          ' their network card's speed in Mb, "" if not known
Private mWhoClickRow As Long          ' the row of the Chat page's list clicked last

' ---- people who have left today (they show greyed out on Who's Here)
Private mGoneBuild As Collection
Private mGoneCount As Long
Private mGoneNick() As String
Private mGoneSeat() As String
Private mGoneHow() As String          ' home, quit, dropped or break
Private mGoneAt() As String           ' when they left, like 21:15

' ---- the Who's Here page
Private mRowCount As Long
Private mRowKind() As Integer         ' 0 = a group heading, 1 = someone signed in, 2 = someone who left
Private mRowIdx() As Long             ' which person, or which heading (0 = signed in, 1 = left)
Private mColX(0 To 5) As Long         ' where each column starts
Private mPplSort As Integer           ' sort by 0 = name, 1 = seat, 2 = status, 3 = server ping, 4 = your ping
Private mPplDesc As Boolean           ' sort the other way round
Private mPplSel As String             ' the person picked in the list
Private mPplDownRow As Long           ' the row the mouse went down on
Private mHdrDown As Integer           ' the column heading being pressed, -1 = none
Private mCard() As String             ' the WHOIS answer about the person picked
Private mCardNick As String
Private mWhoView As Integer           ' 0 = the list, 1 = the seat map
Private mWhoTick As Long              ' seconds on Who's Here, for refreshing pings
Private mMap As SeatMap               ' the seat map in use, from the server
Private mLay As MapLayout             ' where its seats are drawn
Private mMapH As Long                 ' how tall the map is, in pixels
Private mMapDownSeat As Long          ' the seat the mouse went down on

' ---- our own ping to each other LAN Center (UDP on port 4099)
Private mPeerIp(0 To 63) As String
Private mPeerMs(0 To 63) As Long
Private mPeerAt(0 To 63) As Long      ' when it last answered (GetTickCount)
Private mPeerSent(0 To 63) As Long    ' when we last asked
Private mPeerCount As Long
Private mPeerDirty As Boolean         ' new answers to show
Private mFindBroken As Boolean        ' the UDP listener needs starting again

' ---- the Game Library page
Private mLibRowCount As Long
Private mLibRowKind() As Integer      ' 0 = a group heading, 1 = a game
Private mLibRowId() As String         ' the game's id; for a heading, "here" or "new"
Private mLibHereN As Long             ' games installed on this PC
Private mLibNewN As Long              ' games that aren't
Private mLibColX(0 To 4) As Long
Private mLibSel As String             ' the game picked
Private mLibDownRow As Long
Private mInLibTick As Boolean
Private mLibToastId As String         ' the game a "click to read the extra steps" pop-up is about
Private mLibBadge As String           ' what the sidebar bubble says, like 45%

' ---- the Tournaments page (modCups keeps the tournaments themselves)
Private mCupSel As String             ' the tournament on show (its id)
Private mCupMatch As Long             ' the match picked: round * 100 + match, -1 = none
Private mCupError As String           ' the server's answer when something didn't work
Private mCupGoKey As String           ' what the button on the bar along the bottom does
Private mCupTabX(0 To 7) As Long      ' where each tab is
Private mCupTabW(0 To 7) As Long
Private mCupToast As String           ' the tournament a pop-up is about
Private mCupToastMatch As Long
Private mCupBadge As String           ' what the red bubble on Tournaments says
Private mCupHeard As Boolean          ' the first list has come in (after that, a new tournament pops up)
Private mCupWait As Long              ' tournaments still to come in the first list
Private mCupNew As String             ' a tournament that has just opened sign-ups

' ---- redrawing
Private mDirty As Long                ' DR_ flags waiting to be drawn
Private mWantWho As Boolean           ' ask the server for the who's here list on the next tick
Private mLastNameToast As Long        ' when we last popped up "someone said your name"

' ================================================================ start up

Private Sub Form_Load()
    Dim cmdLine As String
    ReDim mLnTime(0 To MAX_LINES - 1)
    ReDim mLnIcon(0 To MAX_LINES - 1)
    ReDim mLnNick(0 To MAX_LINES - 1)
    ReDim mLnNickCol(0 To MAX_LINES - 1)
    ReDim mLnNickW(0 To MAX_LINES - 1)
    ReDim mLnText(0 To MAX_LINES - 1)
    ReDim mLnCol(0 To MAX_LINES - 1)
    ReDim mLnBg(0 To MAX_LINES - 1)
    ReDim mLnWrap(0 To MAX_LINES - 1)
    ReDim mLnRows(0 To MAX_LINES - 1)
    ReDim mWhoNick(0 To 0)
    ReDim mWhoSeat(0 To 0)
    ReDim mWhoAway(0 To 0)
    ReDim mWhoBusy(0 To 0)
    ReDim mWhoAdmin(0 To 0)
    ReDim mWhoSince(0 To 0)
    ReDim mWhoIp(0 To 0)
    ReDim mWhoMs(0 To 0)
    ReDim mWhoLink(0 To 0)
    ReDim mMap.Tables(0 To 0)
    mMapDownSeat = -1
    mWhoView = Val(SettingGet("who_view", "0"))
    If mWhoView <> 1 Then mWhoView = 0
    ReDim mGoneNick(0 To 0)
    ReDim mGoneSeat(0 To 0)
    ReDim mGoneHow(0 To 0)
    ReDim mGoneAt(0 To 0)
    ReDim mRowKind(0 To 0)
    ReDim mRowIdx(0 To 0)
    ReDim mCard(0 To 0)
    mHdrDown = -1
    mPplDownRow = -1
    mWhoClickRow = -1
    mPplSort = Val(SettingGet("who_sort", "0"))
    If mPplSort < 0 Or mPplSort > 4 Then mPplSort = 0
    mPplDesc = (SettingGet("who_desc", "0") = "1")
    mNavHover = -1
    mNavDown = -1
    mToolHover = -1
    mToolDown = -1
    mPingMs = -1

    ReDim mSchedId(0 To 0)
    ReDim mSchedTime(0 To 0)
    ReDim mSchedIcon(0 To 0)
    ReDim mSchedName(0 To 0)
    ReDim mSchedIn(0 To 0)
    ReDim mSchedInN(0 To 0)
    ReDim mSchedNote(0 To 0)
    mSchedSel = -1
    mCupMatch = -1
    mVoteState = "none"
    mRemind = SettingGet("remind", "")
    mSuggFresh = True
    mFoodState = "closed"

    ' the Game Library: what's installed on this PC
    ReDim mLibRowKind(0 To 0)
    ReDim mLibRowId(0 To 0)
    mLibDownRow = -1
    mLibSel = SettingGet("lib_sel", "")
    ' an older version remembered a pak, like quake: now it's a game in a pak, quake/quake
    If Len(mLibSel) > 0 And InStr(mLibSel, "/") = 0 Then mLibSel = mLibSel & "/" & mLibSel
    ReDim gPak(0 To 0)
    ReDim gGame(0 To 0)
    LoadHere
    LibRebuild
    BuildLibRows

    LoadGfx
    DrawSmiley False

    mnuToolsSounds.Checked = gSoundOn

    ' Each PC gets its own ID. If these settings came from another PC (a copied disk image),
    ' start fresh, or the two PCs would keep signing each other out.
    mPcId = SettingGet("pcid", "")
    If Len(SettingGet("pcname", "")) > 0 And SettingGet("pcname", "") <> ThisPCName() Then
        mPcId = ""
        SettingSet "nick", ""
        SettingSet "seat", ""
    End If
    If Len(mPcId) = 0 Then
        mPcId = NewId()
        SettingSet "pcid", mPcId
    End If
    SettingSet "pcname", ThisPCName()
    txtNick.Text = SettingGet("nick", "")
    txtSeat.Text = SettingGet("seat", "")
    mServerPort = Val(SettingGet("port", CStr(DEFAULT_PORT)))
    If mServerPort <= 0 Then mServerPort = DEFAULT_PORT
    mServerHost = SettingGet("server", "")
    txtServer.Text = mServerHost

    Randomize
    cmdLine = LCase$(Command$)
    If InStr(cmdLine, "/min") > 0 Then Me.WindowState = vbMinimized
    If InStr(cmdLine, "/signin") > 0 And Len(SettingGet("nick", "")) > 0 Then
        mMyNick = SettingGet("nick", "")
        mMySeat = SettingGet("seat", "")
        mAutoRelogin = True
    End If

    ShowPage PG_SIGNIN
    WheelHook Me.hWnd
    StartListening
    If Len(mServerHost) > 0 Then
        ConnectTo mServerHost
    Else
        SetConn IC_REFRESH, "Searching for the Palms Server..."
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    LayoutAll
End Sub

' Closing with a game still installing: check first
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then Exit Sub
    If Len(gJobId) > 0 Then
        If MsgBox(LibPakName(gJobId) & " is being installed." & vbCrLf & vbCrLf & _
                  "Cancel the installation and exit? Files already downloaded are kept, and the download resumes next time.", vbYesNo + vbQuestion + vbDefaultButton2, "Palms LAN Center") <> vbYes Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If CartBuilding() Then
        If MsgBox("A cartridge is being made." & vbCrLf & vbCrLf & _
                  "Stop making it and exit? The incomplete cartridge will be removed.", vbYesNo + vbQuestion + vbDefaultButton2, "Palms LAN Center") <> vbYes Then Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, t As Single
    On Error Resume Next
    tmrLib.Enabled = False
    LibShutdown
    WheelUnhook
    tmrClock.Enabled = False
    mClosing = True
    If sckTCP.State = sckConnected Then
        SendLine "QUIT"
        ' give the goodbye half a second to reach the server
        t = Timer
        Do While Timer - t < 0.5 And Timer >= t
            DoEvents
        Loop
    End If
    tmrPing.Enabled = False
    tmrReconnect.Enabled = False
    tmrHover.Enabled = False
    tmrNote.Enabled = False
    tmrLazy.Enabled = False
    sckTCP.Close
    sckFind.Close
    For i = Forms.Count - 1 To 0 Step -1
        If Not Forms(i) Is Me Then Unload Forms(i)
    Next
End Sub

' ================================================================ laying things out

Private Sub LayoutAll()
    Dim w As Long, h As Long, yTop As Long, bodyH As Long
    w = Me.ScaleWidth
    h = Me.ScaleHeight
    If w < 560 Then w = 560
    If h < 380 Then h = 380
    picTools.Move 0, 0, w, 30
    picStatus.Move 0, h - 22, w, 22
    yTop = 32
    bodyH = h - 22 - yTop - 2
    picSide.Move 2, yTop, 88, bodyH
    picPane.Move 2 + 88 + 4, yTop, w - (2 + 88 + 4) - 2, bodyH
    LayoutPane
    DrawTools
    DrawSide
    DrawStatus
End Sub

Private Sub LayoutPane()
    Dim w As Long, h As Long, pgTop As Long, pgH As Long
    w = picPane.ScaleWidth
    h = picPane.ScaleHeight
    picPane.Cls
    Sunken2 picPane, 0, 0, w, h
    picHeader.Move 2, 2, w - 4, 53
    pgTop = 2 + 53
    pgH = h - pgTop - 2
    picSignInPage.Move 2, pgTop, w - 4, pgH
    picChatPage.Move 2, pgTop, w - 4, pgH
    picPAPage.Move 2, pgTop, w - 4, pgH
    picSoonPage.Move 2, pgTop, w - 4, pgH
    picWhatsPage.Move 2, pgTop, w - 4, pgH
    picWhoPage.Move 2, pgTop, w - 4, pgH
    picLibPage.Move 2, pgTop, w - 4, pgH
    picCupsPage.Move 2, pgTop, w - 4, pgH
    LayoutSignIn
    LayoutChat
    LayoutPA
    LayoutWhats
    LayoutWho
    LayoutLibrary
    LayoutCups
    DrawHeader
    DrawSoon
End Sub

Private Sub LayoutSignIn()
    Dim w As Long, gx As Long, gy As Long
    w = picSignInPage.ScaleWidth
    gx = (w - 400) \ 2
    If gx < 10 Then gx = 10
    gy = 24
    picSignInPage.Cls
    GroupBox picSignInPage, gx, gy, 400, 196, IC_USER, "Sign In"
    lblFind.Move gx + 16, gy + 18, 368, 30
    lblNick.Move gx + 16, gy + 60, 110, 16
    txtNick.Move gx + 130, gy + 57, 150, 21
    lblSeat.Move gx + 16, gy + 89, 110, 16
    txtSeat.Move gx + 130, gy + 86, 50, 21
    lblServer.Move gx + 16, gy + 118, 110, 16
    txtServer.Move gx + 130, gy + 115, 150, 21
    cmdSignIn.Move gx + 400 - 16 - 90 - 8 - 90, gy + 156, 90, 25
    cmdQuit.Move gx + 400 - 16 - 90, gy + 156, 90, 25
    TipBox picSignInPage, gx, gy + 212, 400, 52
    DrawIcon16 picSignInPage, IC_INFO, gx + 8, gy + 220
    TextAt picSignInPage, gx + 30, gy + 221, "Nicknames are 2 to 12 characters long.", vbBlack
    TextAt picSignInPage, gx + 30, gy + 239, "The seat number is optional. It can be set after signing in.", vbBlack
End Sub

Private Sub LayoutChat()
    Dim w As Long, h As Long, rx As Long, lw As Long, nh As Long, gTop As Long, gBot As Long
    w = picChatPage.ScaleWidth
    h = picChatPage.ScaleHeight
    rx = w - 10 - 178
    lw = rx - 10 - 10
    If lw < 150 Then lw = 150
    picLog.Move 10, 10, lw, h - 10 - 23 - 6 - 10
    txtSay.Move 10, h - 10 - 23, lw - 80 - 4 - 23 - 4, 23
    picSmiley.Move 10 + lw - 80 - 4 - 23, h - 10 - 23, 23, 23
    DrawSmiley False
    cmdSend.Move 10 + lw - 80, h - 10 - 23, 80, 23
    nh = 74
    gTop = 16
    gBot = h - 10 - nh - 10
    picNext.Move rx, h - 10 - nh, 178, nh
    picWho.Move rx + 6, gTop + 12, 178 - 12, gBot - 6 - (gTop + 12)
    picChatPage.Cls
    DrawWhoGroup
    vsbLog.Move picLog.ScaleWidth - 16, 0, 16, picLog.ScaleHeight
    vsbWho.Move picWho.ScaleWidth - 16, 0, 16, picWho.ScaleHeight
    WrapAll
    UpdateLogScroll (vsbLog.Value >= vsbLog.Max)
    DrawLog
    UpdateWhoScroll
    DrawWho
    DrawNext
End Sub

Private Sub DrawWhoGroup()
    Want DR_WHO
End Sub

Private Sub DrawWhoGroupNow()
    Dim w As Long, h As Long, gBot As Long
    w = picChatPage.ScaleWidth
    h = picChatPage.ScaleHeight
    gBot = h - 10 - 74 - 10
    Box picChatPage, w - 10 - 178, 4, 178, gBot - 4, vbButtonFace
    GroupBox picChatPage, w - 10 - 178, 16, 178, gBot - 16, IC_PEOPLE, "Who's Here (" & mWhoCount & ")"
End Sub

Private Sub LayoutPA()
    Dim w As Long
    w = picPAPage.ScaleWidth
    picPAPage.Cls
    lblPA.Move 10, 10, w - 20, 16
    txtPA.Move 10, 28, w - 20, 58
    GroupBox picPAPage, 10, 104, w - 20, 62, IC_SPEAKER, "Chime"
    optPA(0).Move 24, 120, 140, 18
    optPA(1).Move 180, 120, 140, 18
    optPA(2).Move 24, 141, 140, 18
    optPA(3).Move 180, 141, 140, 18
    cmdPATest.Move 10, 178, 110, 25
    cmdPASend.Move 126, 178, 120, 25
    TipBox picPAPage, 10, 216, w - 20, 52
    DrawIcon16 picPAPage, IC_INFO, 18, 224
    TextAt picPAPage, 40, 225, "The announcement is shown on every PC signed in, after the chime.", vbBlack
    TextAt picPAPage, 40, 243, "Announcements can be up to 200 characters long.", vbBlack
End Sub

' ================================================================ pages

Private Function PageName(ByVal pg As Integer) As String
    Select Case pg
    Case PG_CHAT: PageName = "Chat"
    Case PG_WHATSON: PageName = "What's On"
    Case PG_PLAY: PageName = "Play Now"
    Case PG_LIBRARY: PageName = "Game Library"
    Case PG_WHO: PageName = "Who's Here"
    Case PG_CUPS: PageName = "Tournaments"
    Case PG_JUKEBOX: PageName = "Jukebox"
    Case PG_PA: PageName = "Announce"
    Case Else: PageName = "Palms LAN Center"
    End Select
End Function

Private Function PageIcon(ByVal pg As Integer) As Integer
    Select Case pg
    Case PG_CHAT: PageIcon = IC_CHAT
    Case PG_WHATSON: PageIcon = IC_CALENDAR
    Case PG_PLAY: PageIcon = IC_JOYSTICK
    Case PG_LIBRARY: PageIcon = IC_DISC
    Case PG_WHO: PageIcon = IC_PEOPLE
    Case PG_CUPS: PageIcon = IC_TROPHY
    Case PG_JUKEBOX: PageIcon = IC_JUKEBOX
    Case PG_PA: PageIcon = IC_MEGAPHONE
    Case Else: PageIcon = IC_PALM
    End Select
End Function

Private Function PageSub(ByVal pg As Integer) As String
    Select Case pg
    Case PG_CHAT: PageSub = "Messages to everyone signed in."
    Case PG_WHATSON: PageSub = "Today's schedule and next game vote."
    Case PG_PLAY: PageSub = "Game servers running on the network."
    Case PG_LIBRARY: PageSub = "Games available from the Cartridge Server. Installed games are marked with a tick."
    Case PG_WHO: PageSub = "Users signed in, their seats and system information."
    Case PG_CUPS: PageSub = "Sign-ups, brackets and results."
    Case PG_JUKEBOX: PageSub = "Music requests and voting."
    Case PG_PA: PageSub = "Send a message with a chime to every PC signed in."
    Case Else: PageSub = "Sign in to the Palms Server to continue."
    End Select
End Function

' The two lines on a "coming soon" page
Private Sub SoonText(ByVal pg As Integer, l1 As String, l2 As String)
    Select Case pg
    Case PG_WHATSON
        l1 = "Shows today's schedule and the next game vote."
        l2 = "Announcements are shown in Chat."
    Case PG_PLAY
        l1 = "Lists game servers running on the network, with a Join button."
        l2 = "Shows each server's address, players and ping."
    Case PG_CUPS
        l1 = "Shows tournament brackets."
        l2 = "Tournaments need Palms Server 0.5.0 or later."
    Case PG_JUKEBOX
        l1 = "Song requests for the venue's sound system."
        l2 = "Request songs, vote on the next song and see what is playing."
    Case Else
        l1 = ""
        l2 = ""
    End Select
End Sub

' Pages that only say "coming soon" (Tournaments, with an older Palms Server)
Private Function IsSoonPage(ByVal pg As Integer) As Boolean
    IsSoonPage = (pg = PG_PLAY Or pg = PG_JUKEBOX Or (pg = PG_CUPS And Not ServerHasCups()))
End Function

' Switch pages when the person asks (sidebar, View menu)
Private Sub GoPage(ByVal pg As Integer)
    If Not mSignedIn Then Exit Sub
    If pg = PG_PA And Not mIsAdmin Then Exit Sub
    ShowPage pg
End Sub

Private Sub ShowPage(ByVal pg As Integer)
    Dim k As Integer
    mPage = pg
    NavShowPage pg
    If pg = PG_CHAT Then mUnread = 0
    picSignInPage.Visible = (pg = PG_SIGNIN)
    picChatPage.Visible = (pg = PG_CHAT)
    picPAPage.Visible = (pg = PG_PA)
    picWhatsPage.Visible = (pg = PG_WHATSON)
    picSoonPage.Visible = IsSoonPage(pg)
    picWhoPage.Visible = (pg = PG_WHO)
    picLibPage.Visible = (pg = PG_LIBRARY)
    picCupsPage.Visible = (pg = PG_CUPS And ServerHasCups())
    For k = 0 To PG_COUNT - 1
        mnuGo(k).Checked = (k = pg)
    Next
    cmdSignIn.Default = False
    cmdSend.Default = False
    cmdPASend.Default = False
    Select Case pg
    Case PG_SIGNIN: cmdSignIn.Default = True
    Case PG_CHAT: cmdSend.Default = True
    Case PG_PA: cmdPASend.Default = True
    End Select
    If pg = PG_SIGNIN Then
        Me.Caption = "Palms LAN Center"
    Else
        Me.Caption = PageName(pg) & " - Palms LAN Center"
    End If
    BuildTools
    DrawTools
    DrawSide
    DrawHeader
    DrawSoon
    If pg = PG_WHATSON Then DrawWhats
    If pg = PG_WHO Then
        ' the first time, show your own card
        If Len(mPplSel) = 0 Then mPplSel = mMyNick
        picPeople.Visible = (mWhoView = 0)
        picMap.Visible = (mWhoView = 1)
        DrawWhoPage
        AskWhois
        PingPeers
    End If
    If pg = PG_LIBRARY Then
        ' the first time, pick the first game
        If Not LibShown(mLibSel) Then
            For k = 0 To mLibRowCount - 1
                If mLibRowKind(k) = 1 And Not LibShown(mLibSel) Then mLibSel = mLibRowId(k)
            Next
            BuildTools
        End If
        DrawLibPage
    End If
    If pg = PG_CUPS Then
        If CupK() < 0 And gCupCount > 0 Then SelectCup gCupId(0)
        UpdateCupScroll
        DrawCupPage
    End If
    UpdateMenus
    On Error Resume Next
    Select Case pg
    Case PG_SIGNIN: txtNick.SetFocus
    Case PG_CHAT: txtSay.SetFocus
    Case PG_PA: txtPA.SetFocus
    Case PG_WHO
        If mWhoView = 1 Then
            picMap.SetFocus
        Else
            picPeople.SetFocus
        End If
    Case PG_LIBRARY
        picLibList.SetFocus
    Case PG_CUPS
        If picCup.Visible Then picCup.SetFocus
    End Select
    ' draw the new page now, along with anything it missed while it was hidden
    PaintNow
End Sub

Private Sub ShowSignIn()
    mSignedIn = False
    mIsAdmin = False
    mAway = False
    mBusy = False
    cmdSignIn.Enabled = True
    ShowPage PG_SIGNIN
End Sub

' ================================================================ drawing: header, coming soon, status

Private Sub DrawHeader()
    Want DR_HEADER
End Sub

Private Sub DrawHeaderNow()
    Dim w As Long, rt As String, ric As Integer, tw As Long
    picHeader.Cls
    w = picHeader.ScaleWidth
    DrawIcon32 picHeader, PageIcon(mPage), 12, 9
    picHeader.FontSize = 14
    TextAt picHeader, 52, 4, PageName(mPage), C_TITLE, True
    picHeader.FontSize = 8
    TextAt picHeader, 53, 31, PageSub(mPage), C_SUB
    rt = ""
    Select Case mPage
    Case PG_CHAT
        ric = IC_PEOPLE
        If mWhoCount = 1 Then
            rt = "1 user signed in"
        Else
            rt = mWhoCount & " users signed in"
        End If
    Case PG_PA
        ric = IC_NET
        rt = "Sent to all users signed in"
    Case PG_WHATSON
        ric = IC_CLOCK
        rt = "Local time: " & AmPm(Format$(Now, "hh:nn"))
    Case PG_WHO
        ric = IC_PEOPLE
        rt = HereText()
    Case PG_LIBRARY
        ric = IC_DISC
        If gShowCount = 1 Then
            rt = "1 game, "
        Else
            rt = gShowCount & " games, "
        End If
        rt = rt & gHereCount & " installed"
    Case PG_CUPS
        If ServerHasCups() Then
            ric = IC_TROPHY
            If gCupCount = 1 Then
                rt = "1 tournament"
            Else
                rt = gCupCount & " tournaments"
            End If
        End If
    End Select
    If Len(rt) > 0 Then
        tw = picHeader.TextWidth(rt)
        DrawIcon16 picHeader, ric, w - 12 - tw - 20, 17
        TextAt picHeader, w - 12 - tw, 19, rt, C_SUB
    End If
    Sunset picHeader, 0, 50, w, 3
End Sub

Private Sub DrawSoon()
    Dim w As Long, l1 As String, l2 As String
    If Not IsSoonPage(mPage) Then Exit Sub
    picSoonPage.Cls
    w = picSoonPage.ScaleWidth
    SoonText mPage, l1, l2
    TipBox picSoonPage, 10, 10, w - 20, 70
    DrawIcon32 picSoonPage, PageIcon(mPage), 22, 29
    TextAt picSoonPage, 66, 22, "Feature coming soon, watch this space.", vbBlack, True
    TextAt picSoonPage, 66, 40, l1, vbBlack
    TextAt picSoonPage, 66, 56, l2, C_SUB
End Sub

Private Sub SetConn(ByVal ic As Integer, ByVal s As String)
    mConnIcon = ic
    mConnText = s
    DrawStatus
End Sub

' A short message in the status bar that goes away after a few seconds
Private Sub Note(ByVal s As String)
    mNoteText = s
    tmrNote.Enabled = False
    tmrNote.Enabled = True
    DrawStatus
End Sub

Private Sub tmrNote_Timer()
    tmrNote.Enabled = False
    mNoteText = ""
    DrawStatus
End Sub

Private Sub DrawStatus()
    Want DR_STATUS
End Sub

Private Sub DrawStatusNow()
    Dim w As Long, x2 As Long, x3 As Long, s As String, ic As Integer
    picStatus.Cls
    w = picStatus.ScaleWidth
    x3 = w - 2 - 132
    x2 = x3 - 2 - 112

    Edge picStatus, 2, 2, x2 - 4, 20, vb3DShadow, vb3DHighlight
    If Len(mNoteText) > 0 Then
        ic = IC_INFO
        s = mNoteText
    Else
        ic = mConnIcon
        s = mConnText
    End If
    DrawIcon16 picStatus, ic, 5, 4
    TextAt picStatus, 25, 5, FitText(picStatus, s, x2 - 4 - 28), vbButtonText

    Edge picStatus, x2, 2, 112, 20, vb3DShadow, vb3DHighlight
    If mSignedIn Then
        s = HereText()
        DrawIcon16 picStatus, IC_PEOPLE, x2 + 3, 4
        TextAt picStatus, x2 + 23, 5, FitText(picStatus, s, 112 - 26), vbButtonText
    End If

    Edge picStatus, x3, 2, 132, 20, vb3DShadow, vb3DHighlight
    If sckTCP.State <> sckConnected Then
        ic = IC_SIG0
        s = "Connection: Offline"
    ElseIf mPingMs < 0 Then
        ic = IC_SIG3
        s = "Connection: Good"
    ElseIf mPingMs < 20 Then
        ic = IC_SIG4
        s = "Connection: Fast"
    ElseIf mPingMs < 80 Then
        ic = IC_SIG3
        s = "Connection: Good"
    Else
        ic = IC_SIG2
        s = "Connection: Slow"
    End If
    DrawIcon16 picStatus, ic, x3 + 3, 4
    TextAt picStatus, x3 + 23, 5, s, vbButtonText
End Sub

' "9 here, 2 away"
Private Function HereText() As String
    Dim i As Long, nHere As Long, nAway As Long, nBusy As Long
    For i = 0 To mWhoCount - 1
        If mWhoAway(i) Then
            nAway = nAway + 1
        ElseIf mWhoBusy(i) Then
            nBusy = nBusy + 1
        Else
            nHere = nHere + 1
        End If
    Next
    HereText = nHere & " online"
    If nAway > 0 Then HereText = HereText & ", " & nAway & " away"
    If nBusy > 0 Then HereText = HereText & ", " & nBusy & " busy"
End Function

' ================================================================ redrawing
' Rather than redrawing straight away every time something changes (a dozen times
' for one burst of messages from the server), we note what needs redrawing and
' draw it all once, a moment later. Pages you can't see wait until you open them.

Private Sub Want(ByVal what As Long)
    mDirty = mDirty Or what
    If Not tmrLazy.Enabled Then tmrLazy.Enabled = True
End Sub

Private Sub tmrLazy_Timer()
    tmrLazy.Enabled = False
    If mWantWho Then
        mWantWho = False
        SendLine "WHO"
    End If
    PaintNow
End Sub

Private Sub PaintNow()
    Dim d As Long, keep As Long
    ' minimised: nothing to see. Form_Resize draws everything when it comes back.
    If Me.WindowState = vbMinimized Then Exit Sub
    d = mDirty
    mDirty = 0
    If (d And DR_TOOLS) <> 0 Then
        PerfStart PF_TOOLS
        DrawToolsNow
        PerfStop PF_TOOLS
    End If
    If (d And DR_SIDE) <> 0 Then
        PerfStart PF_SIDE
        DrawSideNow
        PerfStop PF_SIDE
    End If
    If (d And DR_STATUS) <> 0 Then
        PerfStart PF_STATUS
        DrawStatusNow
        PerfStop PF_STATUS
    End If
    If (d And DR_HEADER) <> 0 Then
        PerfStart PF_HEADER
        DrawHeaderNow
        PerfStop PF_HEADER
    End If
    ' Chat
    If mPage = PG_CHAT Then
        If (d And DR_WHO) <> 0 Then
            PerfStart PF_WHO
            DrawWhoGroupNow
            DrawWhoNow
            PerfStop PF_WHO
        End If
        If (d And DR_LOG) <> 0 Then
            PerfStart PF_LOG
            DrawLogNow
            PerfStop PF_LOG
        End If
        If (d And DR_NEXT) <> 0 Then
            PerfStart PF_NEXT
            DrawNextNow
            PerfStop PF_NEXT
        End If
    Else
        keep = keep Or (d And (DR_WHO Or DR_LOG Or DR_NEXT))
    End If
    ' What's On
    If mPage = PG_WHATSON Then
        If (d And DR_WHATS) <> 0 Then
            PerfStart PF_WHATS
            DrawWhatsNow
            PerfStop PF_WHATS
        Else
            If (d And DR_SCHED) <> 0 Then
                PerfStart PF_SCHED
                DrawSchedNow
                PerfStop PF_SCHED
            End If
            If (d And DR_VOTE) <> 0 Then
                PerfStart PF_VOTE
                DrawVoteNow
                PerfStop PF_VOTE
            End If
        End If
    Else
        keep = keep Or (d And (DR_WHATS Or DR_SCHED Or DR_VOTE))
    End If
    ' Who's Here
    If mPage = PG_WHO Then
        If (d And DR_WHOPAGE) <> 0 Then
            PerfStart PF_WHOPAGE
            DrawWhoPageNow
            PerfStop PF_WHOPAGE
        Else
            If (d And DR_WHOBODY) <> 0 Then
                PerfStart PF_WHOBODY
                If mWhoView = 1 Then
                    DrawMapViewNow
                Else
                    DrawPeopleNow
                End If
                PerfStop PF_WHOBODY
            End If
            If (d And DR_CARD) <> 0 Then
                PerfStart PF_CARD
                DrawCardNow
                PerfStop PF_CARD
            End If
        End If
    Else
        keep = keep Or (d And (DR_WHOPAGE Or DR_WHOBODY Or DR_CARD))
    End If
    ' Game Library
    If mPage = PG_LIBRARY Then
        If (d And DR_LIBPAGE) <> 0 Then
            PerfStart PF_LIBPAGE
            DrawLibPageNow
            PerfStop PF_LIBPAGE
        Else
            If (d And DR_LIBLIST) <> 0 Then
                PerfStart PF_LIBLIST
                DrawLibListNow
                PerfStop PF_LIBLIST
            End If
            If (d And DR_LIBCARD) <> 0 Then
                PerfStart PF_LIBCARD
                DrawLibCardNow
                UpdateLibButtons
                PerfStop PF_LIBCARD
            End If
        End If
    Else
        keep = keep Or (d And (DR_LIBPAGE Or DR_LIBLIST Or DR_LIBCARD))
    End If
    ' Tournaments
    If mPage = PG_CUPS Then
        If (d And DR_CUPPAGE) <> 0 Then
            PerfStart PF_CUPS
            DrawCupPageNow
            PerfStop PF_CUPS
        ElseIf (d And DR_CUP) <> 0 Then
            PerfStart PF_CUPS
            DrawCupNow
            PerfStop PF_CUPS
        End If
    Else
        keep = keep Or (d And (DR_CUPPAGE Or DR_CUP))
    End If
    ' the pages you can't see keep their place in the queue
    mDirty = mDirty Or keep
End Sub

' ================================================================ the sidebar

Private Function NavCount() As Integer
    If mIsAdmin Then
        NavCount = 8
    Else
        NavCount = 7
    End If
End Function

' Every sidebar button is the same height. When they don't all fit, scroll arrows
' appear at the top and bottom of the sidebar, like the Outlook bar.
Private Function NavFits() As Boolean
    NavFits = (NAV_TOP + NavCount() * (NAV_H + NAV_GAP) + 4 <= picSide.ScaleHeight)
End Function

' Where the first button shown goes
Private Function NavY0() As Long
    If NavFits() Then
        NavY0 = NAV_TOP
    Else
        NavY0 = 4 + NAV_ARROW_H + 4
    End If
End Function

' How many buttons fit between the arrows
Private Function NavShown() As Integer
    Dim n As Long
    If NavFits() Then
        NavShown = NavCount()
        Exit Function
    End If
    n = (picSide.ScaleHeight - 2 * (4 + NAV_ARROW_H + 4) + NAV_GAP) \ (NAV_H + NAV_GAP)
    If n < 1 Then n = 1
    If n > NavCount() Then n = NavCount()
    NavShown = n
End Function

' Keeps the first button shown in range
Private Sub NavClamp()
    Dim mx As Integer
    If NavFits() Then
        mNavFirst = 0
        Exit Sub
    End If
    mx = NavCount() - NavShown()
    If mNavFirst > mx Then mNavFirst = mx
    If mNavFirst < 0 Then mNavFirst = 0
End Sub

' Scrolls the sidebar so the button for page pg can be seen
Private Sub NavShowPage(ByVal pg As Integer)
    If pg < 0 Then Exit Sub
    If NavFits() Then Exit Sub
    If pg < mNavFirst Then mNavFirst = pg
    If pg > mNavFirst + NavShown() - 1 Then mNavFirst = pg - NavShown() + 1
    NavClamp
End Sub

Private Sub NavScroll(ByVal delta As Integer)
    mNavFirst = mNavFirst + delta
    NavClamp
    mNavHover = -1
    DrawSide
End Sub

Private Function NavAt(ByVal px As Single, ByVal py As Single) As Integer
    Dim i As Integer, y As Long
    NavAt = -1
    If Not mSignedIn Then Exit Function
    If px < 4 Or px >= picSide.ScaleWidth - 4 Then Exit Function
    NavClamp
    For i = mNavFirst To mNavFirst + NavShown() - 1
        y = NavY0() + (i - mNavFirst) * (NAV_H + NAV_GAP)
        If py >= y And py < y + NAV_H Then NavAt = i
    Next
End Function

' -1 = the up arrow, 1 = the down arrow, 0 = neither
Private Function NavArrowAt(ByVal px As Single, ByVal py As Single) As Integer
    NavArrowAt = 0
    If NavFits() Then Exit Function
    If px < 4 Or px >= picSide.ScaleWidth - 4 Then Exit Function
    If py >= 4 And py < 4 + NAV_ARROW_H Then NavArrowAt = -1
    If py >= picSide.ScaleHeight - 4 - NAV_ARROW_H And py < picSide.ScaleHeight - 4 Then NavArrowAt = 1
End Function

Private Sub DrawSide()
    Want DR_SIDE
End Sub

Private Sub DrawSideNow()
    Dim i As Integer
    picSide.Cls
    Sunken2 picSide, 0, 0, picSide.ScaleWidth, picSide.ScaleHeight
    NavClamp
    For i = mNavFirst To mNavFirst + NavShown() - 1
        DrawNavItem i
    Next
    If Not NavFits() Then
        DrawNavArrow -1
        DrawNavArrow 1
    End If
End Sub

' A scroll arrow: a raised bar the width of the sidebar with a small triangle on it
Private Sub DrawNavArrow(ByVal which As Integer)
    Dim x As Long, y As Long, w As Long, cx As Long, cy As Long, k As Long, col As Long, can As Boolean, d As Long
    x = 4
    w = picSide.ScaleWidth - 8
    If which < 0 Then
        y = 4
        can = (mNavFirst > 0)
    Else
        y = picSide.ScaleHeight - 4 - NAV_ARROW_H
        can = (mNavFirst + NavShown() < NavCount())
    End If
    Box picSide, x, y, w, NAV_ARROW_H, C_SIDE
    If can And mNavArrowDown = which Then
        Edge picSide, x, y, w, NAV_ARROW_H, C_SIDE_DK, C_SIDE_LT
        d = 1
    Else
        Edge picSide, x, y, w, NAV_ARROW_H, C_SIDE_LT, C_SIDE_DK
    End If
    If can Then
        col = vbWhite
    Else
        col = C_SIDE_DIM
    End If
    cx = x + w \ 2 + d
    cy = y + NAV_ARROW_H \ 2 + d
    For k = 0 To 3
        If which < 0 Then
            Box picSide, cx - k, cy - 2 + k, 2 * k + 1, 1, col
        Else
            Box picSide, cx - k, cy + 1 - k, 2 * k + 1, 1, col
        End If
    Next
End Sub

Private Sub DrawNavItem(ByVal i As Integer)
    Dim x As Long, y As Long, iw As Long, s As String, tw As Long, col As Long
    Dim bx As Long, bw As Long, nh As Long, iy As Long
    nh = NAV_H
    x = 4
    y = NavY0() + (i - mNavFirst) * (NAV_H + NAV_GAP)
    iw = picSide.ScaleWidth - 8
    ' the icon and words, in the middle of the button
    iy = y + (nh - 49) \ 2
    If iy < y + 1 Then iy = y + 1
    If i = mPage Then
        Box picSide, x, y, iw, nh, C_SIDE_ON
        Edge picSide, x, y, iw, nh, C_SIDE_DK, C_SIDE_LT
    ElseIf i = mNavHover Then
        If i = mNavDown Then
            Edge picSide, x, y, iw, nh, C_SIDE_DK, C_SIDE_LT
        Else
            Edge picSide, x, y, iw, nh, C_SIDE_LT, C_SIDE_DK
        End If
    End If
    DrawIcon32 picSide, PageIcon(i), x + (iw - 32) \ 2, iy
    s = PageName(i)
    picSide.FontBold = (i = mPage)
    tw = picSide.TextWidth(s)
    picSide.FontBold = False
    If mSignedIn Then
        col = vbWhite
    Else
        col = C_SIDE_DIM
    End If
    TextAt picSide, x + (iw - tw) \ 2, iy + 32 + 4, s, col, (i = mPage)

    ' a red bubble on Chat when there are messages you haven't seen
    If i = PG_CHAT And mUnread > 0 And mPage <> PG_CHAT Then
        If mUnread > 9 Then
            s = "9+"
        Else
            s = CStr(mUnread)
        End If
        picSide.FontBold = True
        tw = picSide.TextWidth(s)
        picSide.FontBold = False
        bw = tw + 6
        If bw < 15 Then bw = 15
        bx = x + (iw - 32) \ 2 + 32 - 8
        Box picSide, bx, iy - 4, bw, 15, vbWhite
        Box picSide, bx + 1, iy - 3, bw - 2, 13, C_BADGE
        TextAt picSide, bx + (bw - tw) \ 2, iy - 3, s, vbWhite, True
    End If

    ' a red bubble on Tournaments: your match is ready, or (the host) results to confirm
    If i = PG_CUPS And mSignedIn And Len(mCupBadge) > 0 Then
        s = mCupBadge
        picSide.FontBold = True
        tw = picSide.TextWidth(s)
        picSide.FontBold = False
        bw = tw + 6
        If bw < 15 Then bw = 15
        bx = x + (iw - 32) \ 2 + 32 - 8
        Box picSide, bx, iy - 4, bw, 15, vbWhite
        Box picSide, bx + 1, iy - 3, bw - 2, 13, C_BADGE
        TextAt picSide, bx + (bw - tw) \ 2, iy - 3, s, vbWhite, True
    End If

    ' a blue bubble on Game Library while a game installs: how far along it is
    If i = PG_LIBRARY And mSignedIn And Len(mLibBadge) > 0 Then
        s = mLibBadge
        picSide.FontBold = True
        tw = picSide.TextWidth(s)
        picSide.FontBold = False
        bw = tw + 6
        If bw < 15 Then bw = 15
        bx = x + (iw - 32) \ 2 + 32 - 8
        If bx + bw > x + iw - 1 Then bx = x + iw - 1 - bw
        Box picSide, bx, iy - 4, bw, 15, vbWhite
        Box picSide, bx + 1, iy - 3, bw - 2, 13, C_BLUE
        TextAt picSide, bx + (bw - tw) \ 2, iy - 3, s, vbWhite, True
    End If
End Sub

Private Sub picSide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    i = NavAt(X, Y)
    If i <> mNavHover Then
        mNavHover = i
        DrawSide
    End If
    If i >= 0 Then tmrHover.Enabled = True
End Sub

Private Sub picSide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    ' the scroll arrows scroll by one button as soon as they're pressed
    mNavArrowDown = NavArrowAt(X, Y)
    If mNavArrowDown <> 0 Then
        NavScroll mNavArrowDown
        Exit Sub
    End If
    mNavDown = NavAt(X, Y)
    mNavHover = mNavDown
    DrawSide
End Sub

Private Sub picSide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, wasDown As Integer
    If Button <> vbLeftButton Then Exit Sub
    If mNavArrowDown <> 0 Then
        mNavArrowDown = 0
        DrawSide
        Exit Sub
    End If
    i = NavAt(X, Y)
    wasDown = mNavDown
    mNavDown = -1
    If i >= 0 And i = wasDown Then
        GoPage i
    Else
        DrawSide
    End If
End Sub

' Clears the hover glow once the mouse leaves the sidebar or toolbar
Private Sub tmrHover_Timer()
    If mToolHover >= 0 And mToolDown < 0 Then
        If Not MouseIsOver(picTools.hWnd) Then
            mToolHover = -1
            DrawTools
        End If
    End If
    If mNavHover >= 0 And mNavDown < 0 Then
        If Not MouseIsOver(picSide.hWnd) Then
            mNavHover = -1
            DrawSide
        End If
    End If
    If mToolHover < 0 And mNavHover < 0 Then tmrHover.Enabled = False
End Sub

' ================================================================ the toolbar

Private Sub AddTool(ByVal key As String, ByVal ic As Integer, ByVal txt As String, ByVal isOn As Boolean)
    If mToolCount >= MAX_TOOLS Then Exit Sub
    mToolKey(mToolCount) = key
    mToolIcon(mToolCount) = ic
    mToolText(mToolCount) = txt
    mToolOn(mToolCount) = isOn
    mToolCount = mToolCount + 1
End Sub

Private Sub BuildTools()
    mToolCount = 0
    mToolHover = -1
    mToolDown = -1
    If Not mSignedIn Then Exit Sub
    Select Case mPage
    Case PG_CHAT
        AddTool "status", MyStatusIcon(), "Status: " & MyStatusWord(), False
        AddTool "-", -1, "", False
        AddTool "alerts", IC_BELL, "Name Alerts", gNotifyName
        If gSoundOn Then
            AddTool "sounds", IC_SPEAKER, "Sounds: On", False
        Else
            AddTool "sounds", IC_SPEAKER, "Sounds: Off", False
        End If
        AddTool "-", -1, "", False
        AddTool "ignore", IC_MUTE, "Mute Players...", False
        AddTool "clear", IC_ERASER, "Clear", False
    Case PG_PA
        AddTool "pa", IC_MEGAPHONE, "Announce", False
        AddTool "patest", IC_SPEAKER, "Test Sound", False
    Case PG_WHATSON
        BuildWhatsTools
    Case PG_LIBRARY
        BuildLibTools
    Case PG_CUPS
        If ServerHasCups() Then BuildCupTools
    Case PG_WHO
        AddTool "seat", IC_TABLE, "Change Seat...", False
        If CanChatTo(mPplSel) Then AddTool "chatto", IC_CHAT, "Chat with " & mPplSel, False
        If mIsAdmin Then
            AddTool "-", -1, "", False
            AddTool "mapedit", IC_PLAN, "Seat Map...", False
        End If
    End Select
End Sub

Private Sub DrawTools()
    Want DR_TOOLS
End Sub

Private Sub DrawToolsNow()
    Dim i As Integer, x As Long, w As Long, lastR As Long, limit As Long, need As Long, bw As Long, full As Boolean
    Dim a As String, b As String, c As String, wa As Long, wb As Long, wc As Long, ic As Integer
    picTools.Cls
    w = picTools.ScaleWidth
    Box picTools, 0, 0, w, 1, vb3DShadow
    Box picTools, 0, 1, w, 1, vb3DHighlight
    ' who you're signed in as goes on the right: work out the room it needs first
    If mSignedIn Then
        a = "Signed in as "
        b = mMyNick
        c = ""
        If Len(mMySeat) > 0 Then c = " at " & mMySeat
        If mAway Then c = c & " (Away)"
        If mBusy Then c = c & " (Busy)"
        If mIsAdmin Then c = c & " (Host)"
        wa = picTools.TextWidth(a)
        wc = picTools.TextWidth(c)
        picTools.FontBold = True
        wb = picTools.TextWidth(b)
        picTools.FontBold = False
    End If
    ' the buttons, left to right. Any that don't fit go on the >> menu.
    need = 3
    For i = 0 To mToolCount - 1
        need = need + ToolWidth(i)
    Next
    If mSignedIn Then
        limit = w - 6 - (6 + 16 + 4 + wb + wc + 4 + 16 + 4) - 12
    Else
        limit = w - 6
    End If
    mChevL = 0
    mChevR = 0
    full = False
    x = 3
    For i = 0 To mToolCount - 1
        bw = ToolWidth(i)
        If Not full And need > limit Then
            If x + bw > limit - CHEV_W Then full = True
        End If
        mToolHide(i) = full
        If full Then
            mToolL(i) = 0
            mToolR(i) = 0
        ElseIf mToolKey(i) = "-" Then
            mToolL(i) = x
            mToolR(i) = x + 8
            Box picTools, x + 3, 5, 1, 20, vb3DShadow
            Box picTools, x + 4, 5, 1, 20, vb3DHighlight
            x = x + 8
        Else
            mToolL(i) = x
            mToolR(i) = x + bw - 1
            DrawToolButton i
            x = mToolR(i) + 1
        End If
    Next
    If full Then
        mChevL = x
        mChevR = x + CHEV_W
        DrawChevron
        x = mChevR + 1
    End If
    lastR = x

    ' who you're signed in as, on the right. Click it for the "Me" menu.
    mMeL = 0
    mMeR = 0
    If mSignedIn Then
        ' if it doesn't fit, drop the "Signed in as" words
        If w - 6 - (6 + 16 + 4 + wa + wb + wc + 4 + 16 + 4) < lastR + 12 Then
            a = ""
            wa = 0
        End If
        mMeR = w - 6
        mMeL = mMeR - (6 + 16 + 4 + wa + wb + wc + 4 + 16 + 4)
        If mMeL > lastR + 4 Then
            If mMeOpen Or (mToolDown = ME_TOOL And mToolHover = ME_TOOL) Then
                Edge picTools, mMeL, 3, mMeR - mMeL, 24, vb3DShadow, vb3DHighlight
            ElseIf mToolHover = ME_TOOL Then
                Edge picTools, mMeL, 3, mMeR - mMeL, 24, vb3DHighlight, vb3DShadow
            End If
            If mAway Then
                ic = IC_MOON
            ElseIf mBusy Then
                ic = IC_BAD
            ElseIf mIsAdmin Then
                ic = IC_ADMIN
            Else
                ic = IC_USER
            End If
            x = mMeL + 6
            DrawIcon16 picTools, ic, x, 7
            x = x + 16 + 4
            TextAt picTools, x, 9, a, C_SUB
            TextAt picTools, x + wa, 9, b, vbButtonText, True
            TextAt picTools, x + wa + wb, 9, c, C_SUB
            DrawIcon16 picTools, IC_DOWN, mMeR - 4 - 16, 7
        Else
            mMeL = 0
            mMeR = 0
        End If
    End If
End Sub

' How much room a toolbar button takes, with the gap after it
Private Function ToolWidth(ByVal i As Integer) As Long
    If mToolKey(i) = "-" Then
        ToolWidth = 8
    Else
        ToolWidth = 6 + 16 + 4 + picTools.TextWidth(mToolText(i)) + 6 + 1
    End If
End Function

' The >> button at the end of the toolbar, when some buttons don't fit
Private Sub DrawChevron()
    Dim w As Long, s As String
    w = mChevR - mChevL
    If mToolDown = CHEV_TOOL And mToolHover = CHEV_TOOL Then
        Edge picTools, mChevL, 3, w, 24, vb3DShadow, vb3DHighlight
    ElseIf mToolHover = CHEV_TOOL Then
        Edge picTools, mChevL, 3, w, 24, vb3DHighlight, vb3DShadow
    End If
    s = Chr$(187)
    picTools.FontBold = True
    TextAt picTools, mChevL + (w - picTools.TextWidth(s)) \ 2, 9, s, vbButtonText, True
    picTools.FontBold = False
End Sub

' Where a button's drop-down menu goes: under the button, or under >> if it's on that menu
Private Function ToolPopX(ByVal i As Integer) As Long
    If mToolHide(i) Then
        ToolPopX = mChevL
    Else
        ToolPopX = mToolL(i)
    End If
End Function

' The >> menu: the buttons that don't fit on the toolbar
Private Sub ShowMoreMenu()
    Dim i As Integer, n As Integer, lastSep As Boolean
    For i = mnuMoreItem.UBound To 1 Step -1
        Unload mnuMoreItem(i)
    Next
    n = 0
    lastSep = True
    For i = 0 To mToolCount - 1
        If mToolHide(i) Then
            If mToolKey(i) = "-" Then
                If Not lastSep Then
                    AddMoreItem n, "-", False, -1
                    lastSep = True
                End If
            Else
                AddMoreItem n, mToolText(i), mToolOn(i), i
                lastSep = False
            End If
        End If
    Next
    If n = 0 Then Exit Sub
    ' no separator at the very end
    If lastSep And n > 1 Then
        Unload mnuMoreItem(n - 1)
        n = n - 1
    End If
    PopupMenu mnuMore, , picTools.Left + mChevL, picTools.Top + 27
End Sub

Private Sub AddMoreItem(n As Integer, ByVal txt As String, ByVal isOn As Boolean, ByVal tool As Integer)
    If n > 23 Then Exit Sub
    If n > 0 Then Load mnuMoreItem(n)
    mnuMoreItem(n).Caption = Replace(txt, "&", "&&")
    mnuMoreItem(n).Checked = isOn
    mnuMoreItem(n).Visible = True
    mMoreMap(n) = tool
    n = n + 1
End Sub

Private Sub mnuMoreItem_Click(Index As Integer)
    If Index < 0 Or Index > 23 Then Exit Sub
    If mMoreMap(Index) >= 0 And mMoreMap(Index) < mToolCount Then ToolClick mMoreMap(Index)
End Sub

Private Sub DrawToolButton(ByVal i As Integer)
    Dim x As Long, w As Long, d As Long
    x = mToolL(i)
    w = mToolR(i) - mToolL(i)
    d = 0
    If mToolOn(i) Then Checker picTools, x, 3, w, 24
    If mToolOn(i) Or (mToolDown = i And mToolHover = i) Then
        Edge picTools, x, 3, w, 24, vb3DShadow, vb3DHighlight
        d = 1
    ElseIf mToolHover = i Then
        Edge picTools, x, 3, w, 24, vb3DHighlight, vb3DShadow
    End If
    DrawIcon16 picTools, mToolIcon(i), x + 6 + d, 7 + d
    TextAt picTools, x + 6 + 16 + 4 + d, 9 + d, mToolText(i), vbButtonText
End Sub

Private Function ToolAt(ByVal px As Single, ByVal py As Single) As Integer
    Dim i As Integer
    ToolAt = -1
    If py < 3 Or py >= 27 Then Exit Function
    If mMeR > mMeL And px >= mMeL And px < mMeR Then
        ToolAt = ME_TOOL
        Exit Function
    End If
    If mChevR > mChevL And px >= mChevL And px < mChevR Then
        ToolAt = CHEV_TOOL
        Exit Function
    End If
    For i = 0 To mToolCount - 1
        If mToolKey(i) <> "-" And Not mToolHide(i) Then
            If px >= mToolL(i) And px < mToolR(i) Then ToolAt = i
        End If
    Next
End Function

Private Sub picTools_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    i = ToolAt(X, Y)
    If i <> mToolHover Then
        mToolHover = i
        DrawTools
    End If
    If i >= 0 Then tmrHover.Enabled = True
End Sub

Private Sub picTools_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    mToolDown = ToolAt(X, Y)
    mToolHover = mToolDown
    DrawTools
End Sub

Private Sub picTools_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, wasDown As Integer
    If Button <> vbLeftButton Then Exit Sub
    i = ToolAt(X, Y)
    wasDown = mToolDown
    mToolDown = -1
    DrawTools
    If i >= 0 And i = wasDown Then ToolClick i
End Sub

Private Sub ToolClick(ByVal i As Integer)
    If i = ME_TOOL Then
        ShowMeMenu
        Exit Sub
    End If
    If i = CHEV_TOOL Then
        ShowMoreMenu
        Exit Sub
    End If
    Select Case mToolKey(i)
    Case "smileys"
        PopupMenu mnuSmileys, , picTools.Left + ToolPopX(i), picTools.Top + 27
    Case "alerts"
        gNotifyName = Not gNotifyName
        SaveSettings
        If gNotifyName Then
            Note "Name alerts enabled."
        Else
            Note "Name alerts disabled."
        End If
        BuildTools
        DrawTools
    Case "sounds"
        ToggleSounds
    Case "status"
        ShowStatusMenu i
    Case "ignore"
        frmIgnore.Show vbModal, Me
    Case "clear"
        ClearLog
        Note "Chat cleared."
    Case "copy"
        CopyChat
    Case "save"
        SaveChat
    Case "pa"
        SendPA
    Case "patest"
        PlayChime PAChime()
    Case "sadd"
        frmEvent.Show vbModal, Me
    Case "tpl"
        frmTemplates.Show vbModal, Me
    Case "sdel"
        RemoveSelectedEvent
    Case "vstart"
        frmVote.Show vbModal, Me
    Case "vstop"
        SendLine "VOTESTOP"
    Case "seat"
        ShowSeatDialog
    Case "chatto"
        ChatTo mPplSel
    Case "mapedit"
        ShowMapEditor
    Case "lgo"
        LibGo
    Case "lupd"
        LibAlt
    Case "lcfg"
        LibPlay mLibSel, Me.hWnd, True
    Case "lnotes"
        LibShowNotes mLibSel
        DrawLibCard
    Case "lfolder"
        LibOpenFolder mLibSel, Me.hWnd
    Case "ldel"
        LibDel
    Case "lroot"
        frmSettings.Show vbModal, Me
    Case "cart"
        ShowCartWizard
    Case "next"
        NextGameGo
    Case "evin"
        ToggleIn mSchedSel
    Case "remind"
        ToggleRemind mSchedSel
    Case "evinfo"
        ShowEventInfo mSchedSel
    Case "evchat"
        EventToChat mSchedSel
    Case "suggest"
        frmSuggest.Show vbModal, Me
    Case "food"
        frmFood.Show vbModal, Me
    Case "print"
        PrintSchedule
    Case "cupjoin", "cupleave", "cupreport", "cupresult", "cupmine", "cupchat", "cupnew", "cupedit", _
         "cupplayers", "cupdraw", "cupredraw", "cupstart", "cupopen", "cupdel"
        CupAction mToolKey(i)
    End Select
End Sub

' The little menu under "Signed in as": your status, and signing out
Private Sub ShowMeMenu()
    If Not mSignedIn Then Exit Sub
    mnuMeHere.Checked = Not mAway And Not mBusy
    mnuMeBreak.Checked = mAway
    mnuMeBusy.Checked = mBusy
    mnuMeBusy.Enabled = ServerHasBusy()
    If Len(mMySeat) > 0 Then
        mnuMeSeat.Caption = "Change &Seat (" & mMySeat & ")..."
    Else
        mnuMeSeat.Caption = "Select &Seat..."
    End If
    mMeOpen = True
    DrawToolsNow
    PopupMenu mnuMe, vbPopupMenuRightAlign, picTools.Left + mMeR, picTools.Top + 27
    mMeOpen = False
    mToolHover = -1
    DrawTools
End Sub

Private Sub mnuMeHere_Click()
    SetMyStatus "here"
End Sub

Private Sub mnuMeBreak_Click()
    SetMyStatus "away"
End Sub

Private Sub mnuMeBusy_Click()
    SetMyStatus "busy"
End Sub

Private Sub mnuStatOnline_Click()
    SetMyStatus "here"
End Sub

Private Sub mnuStatAway_Click()
    SetMyStatus "away"
End Sub

Private Sub mnuStatBusy_Click()
    SetMyStatus "busy"
End Sub

' The Status button on the Chat toolbar
Private Sub ShowStatusMenu(ByVal i As Integer)
    mnuStatOnline.Checked = Not mAway And Not mBusy
    mnuStatAway.Checked = mAway
    mnuStatBusy.Checked = mBusy
    mnuStatBusy.Enabled = ServerHasBusy()
    PopupMenu mnuStat, , picTools.Left + ToolPopX(i), picTools.Top + 27
End Sub

' Online, Away or Busy. The server tells everyone, us included, and STATUS updates the toolbar.
Private Sub SetMyStatus(ByVal how As String)
    Select Case how
    Case "here"
        If mAway Or mBusy Then SendLine "BACK"
    Case "away"
        If Not mAway Then SendLine "AWAY"
    Case "busy"
        If Not mBusy And ServerHasBusy() Then SendLine "BUSY"
    End Select
End Sub

Private Function MyStatusWord() As String
    If mBusy Then
        MyStatusWord = "Busy"
    ElseIf mAway Then
        MyStatusWord = "Away"
    Else
        MyStatusWord = "Online"
    End If
End Function

Private Function MyStatusIcon() As Integer
    If mBusy Then
        MyStatusIcon = IC_BAD
    ElseIf mAway Then
        MyStatusIcon = IC_MOON
    Else
        MyStatusIcon = IC_USER
    End If
End Function

' Servers from 0.4.6 know about Busy
Private Function ServerHasBusy() As Boolean
    ServerHasBusy = (VerNum(mServerVer) >= 406)
End Function

' While Busy, only announcements from the host make a sound or pop up
Private Sub Ding()
    If Not mBusy Then PlayChime "ding"
End Sub

Private Sub mnuMeSeat_Click()
    ShowSeatDialog
End Sub

Private Sub ShowSeatDialog()
    If Not mSignedIn Then Exit Sub
    frmSeat.Show vbModal, Me
End Sub

Private Sub mnuMeSignOut_Click()
    If MsgBox("Sign out? Your seat will be released.", vbQuestion + vbOKCancel, "Sign Out") = vbOK Then
        mSigningOut = True
        mAutoRelogin = False
        SendLine "LOGOUT|home"
    End If
End Sub

Private Sub ToggleSounds()
    gSoundOn = Not gSoundOn
    mnuToolsSounds.Checked = gSoundOn
    SaveSettings
    If gSoundOn Then
        Note "Sounds enabled."
    Else
        Note "Sounds disabled. Announcements are still shown."
    End If
    BuildTools
    DrawTools
End Sub

' ================================================================ the chat log

Private Function LogVisibleRows() As Long
    LogVisibleRows = (picLog.ScaleHeight - 6) \ ROW_H
    If LogVisibleRows < 1 Then LogVisibleRows = 1
End Function

' Breaks text into rows that fit. Rows are joined with a line feed.
Private Function WrapRows(ByVal firstAvail As Long, ByVal fullAvail As Long, ByVal s As String) As String
    Dim words() As String, i As Long, cur As String, avail As Long, out As String, trial As String
    picLog.FontBold = False
    words = Split(s, " ")
    avail = firstAvail
    For i = 0 To UBound(words)
        If Len(cur) = 0 Then
            trial = words(i)
        Else
            trial = cur & " " & words(i)
        End If
        If Len(cur) = 0 Or picLog.TextWidth(trial) <= avail Then
            cur = trial
        Else
            out = out & cur & vbLf
            cur = words(i)
            avail = fullAvail
        End If
    Next
    WrapRows = out & cur
End Function

Private Sub WrapLine(ByVal i As Long)
    Dim full As Long
    full = picLog.ScaleWidth - 16 - TEXT_X - 6
    If full < 60 Then full = 60
    mLnWrap(i) = WrapRows(full - mLnNickW(i), full, mLnText(i))
    mLnRows(i) = UBound(Split(mLnWrap(i), vbLf)) + 1
    If mLnRows(i) < 1 Then mLnRows(i) = 1
End Sub

Private Sub WrapAll()
    Dim i As Long
    For i = 0 To mLineCount - 1
        WrapLine i
    Next
End Sub

Private Sub UpdateLogScroll(ByVal stick As Boolean)
    Dim total As Long, vis As Long, mx As Long, i As Long
    For i = 0 To mLineCount - 1
        total = total + mLnRows(i)
    Next
    vis = LogVisibleRows()
    mx = total - vis
    If mx < 0 Then mx = 0
    If mx > 32000 Then mx = 32000
    mIgnoreScroll = True
    If vsbLog.Value > mx Then vsbLog.Value = mx
    vsbLog.Max = mx
    If vis > 1 Then
        vsbLog.LargeChange = vis - 1
    Else
        vsbLog.LargeChange = 1
    End If
    If stick Then vsbLog.Value = mx
    vsbLog.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

Private Sub DrawLog()
    Want DR_LOG
End Sub

Private Sub DrawLogNow()
    Dim w As Long, vis As Long, topRow As Long, row As Long, i As Long, k As Long
    Dim rows() As String, y As Long, x As Long
    picLog.Cls
    w = picLog.ScaleWidth - 16
    vis = LogVisibleRows()
    topRow = vsbLog.Value
    row = 0
    For i = 0 To mLineCount - 1
        If row + mLnRows(i) > topRow Then
            rows = Split(mLnWrap(i), vbLf)
            If UBound(rows) < 0 Then
                ReDim rows(0 To 0)
            End If
            For k = 0 To UBound(rows)
                If row + k >= topRow And row + k <= topRow + vis Then
                    y = 3 + (row + k - topRow) * ROW_H
                    If mLnBg(i) <> -1 Then Box picLog, 3, y, w - 6, ROW_H, mLnBg(i)
                    If k = 0 Then
                        TextAt picLog, TIME_X, y + 2, mLnTime(i), C_GREY
                        DrawIcon16 picLog, mLnIcon(i), ICON_X, y
                        x = TEXT_X
                        If Len(mLnNick(i)) > 0 Then
                            TextAt picLog, x, y + 2, mLnNick(i) & ":", mLnNickCol(i), True
                            x = x + mLnNickW(i)
                        End If
                        TextAt picLog, x, y + 2, rows(k), mLnCol(i)
                    Else
                        TextAt picLog, TEXT_X, y + 2, rows(k), mLnCol(i)
                    End If
                End If
            Next
        End If
        row = row + mLnRows(i)
        If row > topRow + vis Then Exit For
    Next
End Sub

Private Sub CopyLine(ByVal src As Long, ByVal dst As Long)
    mLnTime(dst) = mLnTime(src)
    mLnIcon(dst) = mLnIcon(src)
    mLnNick(dst) = mLnNick(src)
    mLnNickCol(dst) = mLnNickCol(src)
    mLnNickW(dst) = mLnNickW(src)
    mLnText(dst) = mLnText(src)
    mLnCol(dst) = mLnCol(src)
    mLnBg(dst) = mLnBg(src)
    mLnWrap(dst) = mLnWrap(src)
    mLnRows(dst) = mLnRows(src)
End Sub

' Adds a line to the chat log. nick = "" for messages from the LAN Center itself.
Private Sub AddLog(ByVal tm As String, ByVal ic As Integer, ByVal nick As String, ByVal nickCol As Long, ByVal txt As String, ByVal col As Long, ByVal bg As Long)
    Dim stick As Boolean, i As Long
    stick = (vsbLog.Value >= vsbLog.Max)
    If mLineCount >= MAX_LINES Then
        For i = 50 To MAX_LINES - 1
            CopyLine i, i - 50
        Next
        mLineCount = MAX_LINES - 50
    End If
    i = mLineCount
    mLnTime(i) = tm
    mLnIcon(i) = ic
    mLnNick(i) = nick
    mLnNickCol(i) = nickCol
    mLnText(i) = txt
    mLnCol(i) = col
    mLnBg(i) = bg
    mLnNickW(i) = 0
    If Len(nick) > 0 Then
        picLog.FontBold = True
        mLnNickW(i) = picLog.TextWidth(nick & ": ")
        picLog.FontBold = False
    End If
    WrapLine i
    mLineCount = mLineCount + 1
    UpdateLogScroll stick
    DrawLog
End Sub

Private Sub AddSys(ByVal ic As Integer, ByVal txt As String, ByVal col As Long)
    AddLog ShortTime(Format$(Now, "hh:nn")), ic, "", 0, txt, col, -1
End Sub

Private Sub ClearLog()
    mLineCount = 0
    UpdateLogScroll True
    DrawLog
End Sub

Private Sub ScrollLog(ByVal delta As Long)
    Dim v As Long
    v = vsbLog.Value + delta
    If v < 0 Then v = 0
    If v > vsbLog.Max Then v = vsbLog.Max
    vsbLog.Value = v
End Sub

Private Sub vsbLog_Change()
    If Not mIgnoreScroll Then DrawLog
End Sub

Private Sub vsbLog_Scroll()
    DrawLog
End Sub

Private Sub picLog_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageUp
        Call ScrollLog(-vsbLog.LargeChange)
    Case vbKeyPageDown
        Call ScrollLog(vsbLog.LargeChange)
    Case vbKeyUp
        Call ScrollLog(-1)
    Case vbKeyDown
        Call ScrollLog(1)
    End Select
End Sub

' Page Up and Page Down scroll the chat while you type
Private Sub txtSay_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyPageUp
        Call ScrollLog(-vsbLog.LargeChange)
    Case vbKeyPageDown
        Call ScrollLog(vsbLog.LargeChange)
    End Select
End Sub

Private Function ChatAsText() As String
    Dim i As Long, s As String
    For i = 0 To mLineCount - 1
        s = s & mLnTime(i) & "  "
        If Len(mLnNick(i)) > 0 Then s = s & mLnNick(i) & ": "
        s = s & mLnText(i) & vbCrLf
    Next
    ChatAsText = s
End Function

Private Sub CopyChat()
    Clipboard.Clear
    Clipboard.SetText ChatAsText()
    Note "Chat copied to the Clipboard."
End Sub

Private Sub SaveChat()
    Dim f As Integer, fileName As String
    On Error GoTo Failed
    fileName = AppDir() & "chat-" & Format$(Now, "yyyy-mm-dd") & ".txt"
    f = FreeFile
    Open fileName For Append As #f
    Print #f, ChatAsText();
    Close #f
    Note "Chat saved to " & fileName & "."
    Exit Sub
Failed:
    MsgBox "Unable to save the chat: " & Err.Description, vbExclamation, "Save Chat"
End Sub

' ================================================================ who's here

Private Sub UpdateWhoScroll()
    Dim vis As Long, mx As Long
    vis = (picWho.ScaleHeight - 4) \ ROW_H
    If vis < 1 Then vis = 1
    mx = mWhoCount - vis
    If mx < 0 Then mx = 0
    mIgnoreScroll = True
    If vsbWho.Value > mx Then vsbWho.Value = mx
    vsbWho.Max = mx
    If vis > 1 Then
        vsbWho.LargeChange = vis - 1
    Else
        vsbWho.LargeChange = 1
    End If
    vsbWho.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

Private Sub DrawWho()
    Want DR_WHO
End Sub

Private Sub DrawWhoNow()
    Dim i As Long, y As Long, w As Long, nm As String, col As Long
    picWho.Cls
    w = picWho.ScaleWidth - 16
    For i = vsbWho.Value To mWhoCount - 1
        y = 2 + (i - vsbWho.Value) * ROW_H
        If y > picWho.ScaleHeight Then Exit For
        If mWhoAdmin(i) Then
            DrawIcon16 picWho, IC_ADMIN, 3, y
        Else
            DrawIcon16 picWho, IC_USER, 3, y
        End If
        nm = mWhoNick(i)
        If StrComp(nm, mMyNick, vbTextCompare) = 0 Then nm = nm & " (you)"
        If mWhoAway(i) Then
            col = C_GREY
        ElseIf mWhoAdmin(i) Then
            col = C_PURPLE
        Else
            col = vbBlack
        End If
        TextAt picWho, 23, y + 2, nm, col
        If mWhoAway(i) Then
            DrawIcon16 picWho, IC_MOON, w - 19, y
        ElseIf mWhoBusy(i) Then
            DrawIcon16 picWho, IC_BAD, w - 19, y
        ElseIf Len(mWhoSeat(i)) > 0 Then
            TextAt picWho, w - 4 - picWho.TextWidth(mWhoSeat(i)), y + 2, mWhoSeat(i), C_GREY
        End If
    Next
End Sub

Private Sub vsbWho_Change()
    If Not mIgnoreScroll Then DrawWho
End Sub

' Double-click someone in the Chat page's list to see them on Who's Here
Private Sub picWho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mWhoClickRow = vsbWho.Value + Int((Y - 2) / ROW_H)
End Sub

Private Sub picWho_DblClick()
    If mWhoClickRow < 0 Or mWhoClickRow >= mWhoCount Then Exit Sub
    mPplSel = mWhoNick(mWhoClickRow)
    GoPage PG_WHO
End Sub

Private Sub vsbWho_Scroll()
    DrawWho
End Sub

Private Sub DrawNext()
    Want DR_NEXT
End Sub

Private Sub DrawNextNow()
    Dim s As String, cur As Long, w As Long, x As Long
    picNext.Cls
    w = picNext.ScaleWidth
    TextAt picNext, 7, 6, "Coming Up", vbBlack, True
    DrawIcon16 picNext, IC_CLOCK, 6, 26
    cur = CurrentEvent()
    If mSchedCount = 0 Then
        s = "No events scheduled."
    ElseIf cur + 1 < mSchedCount Then
        s = AmPm(mSchedTime(cur + 1)) & " - " & mSchedName(cur + 1)
    Else
        s = "No more events today."
    End If
    TextAt picNext, 27, 28, FitText(picNext, s, w - 32), vbBlack
    DrawIcon16 picNext, IC_STAR, 6, 47
    Select Case mVoteState
    Case "open"
        s = "Vote closes in " & MmSs(VoteSecsLeft()) & ". "
        TextAt picNext, 27, 49, s, vbBlack
        x = 27 + picNext.TextWidth(s)
        picNext.FontUnderline = True
        TextAt picNext, x, 49, "Vote", C_BLUE
        picNext.FontUnderline = False
    Case "closed"
        If mVoteWinner > 0 Then
            s = mVoteName(mVoteWinner - 1) & " won the vote."
        Else
            s = "Voting closed."
        End If
        TextAt picNext, 27, 49, FitText(picNext, s, w - 32), vbBlack
    Case Else
        TextAt picNext, 27, 49, "No vote in progress.", vbBlack
    End Select
End Sub

' Click the Coming up box to open What's On
Private Sub picNext_Click()
    GoPage PG_WHATSON
End Sub

Private Sub picNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSignedIn Then HandCursor
End Sub

Private Function WhoRank(ByVal i As Long) As String
    Dim r As String
    If StrComp(mWhoNick(i), mMyNick, vbTextCompare) = 0 Then
        r = "0"
    ElseIf mWhoAdmin(i) Then
        r = "1"
    Else
        r = "2"
    End If
    WhoRank = r & LCase$(mWhoNick(i))
End Function

Private Sub SwapWho(ByVal a As Long, ByVal b As Long)
    Dim s As String, f As Boolean
    s = mWhoNick(a)
    mWhoNick(a) = mWhoNick(b)
    mWhoNick(b) = s
    s = mWhoSeat(a)
    mWhoSeat(a) = mWhoSeat(b)
    mWhoSeat(b) = s
    f = mWhoAway(a)
    mWhoAway(a) = mWhoAway(b)
    mWhoAway(b) = f
    f = mWhoBusy(a)
    mWhoBusy(a) = mWhoBusy(b)
    mWhoBusy(b) = f
    f = mWhoAdmin(a)
    mWhoAdmin(a) = mWhoAdmin(b)
    mWhoAdmin(b) = f
    s = mWhoSince(a)
    mWhoSince(a) = mWhoSince(b)
    mWhoSince(b) = s
    s = mWhoIp(a)
    mWhoIp(a) = mWhoIp(b)
    mWhoIp(b) = s
    s = mWhoMs(a)
    mWhoMs(a) = mWhoMs(b)
    mWhoMs(b) = s
    s = mWhoLink(a)
    mWhoLink(a) = mWhoLink(b)
    mWhoLink(b) = s
End Sub

' Turns the WHO lines we collected into the list: you first, then admins, then A to Z
Private Sub FinishWho()
    Dim n As Long, i As Long, j As Long, f() As String, v As Variant
    n = 0
    If Not mWhoBuild Is Nothing Then n = mWhoBuild.Count
    If n > 0 Then
        ReDim mWhoNick(0 To n - 1)
        ReDim mWhoSeat(0 To n - 1)
        ReDim mWhoAway(0 To n - 1)
        ReDim mWhoBusy(0 To n - 1)
        ReDim mWhoAdmin(0 To n - 1)
        ReDim mWhoSince(0 To n - 1)
        ReDim mWhoIp(0 To n - 1)
        ReDim mWhoMs(0 To n - 1)
        ReDim mWhoLink(0 To n - 1)
        i = 0
        For Each v In mWhoBuild
            f = Split(CStr(v), "|")
            mWhoNick(i) = Fld(f, 1)
            mWhoSeat(i) = Fld(f, 2)
            If mWhoSeat(i) = "-" Then mWhoSeat(i) = ""
            mWhoAway(i) = (Fld(f, 3) = "away")
            mWhoBusy(i) = (Fld(f, 3) = "busy")
            mWhoAdmin(i) = (Fld(f, 4) = "1")
            mWhoSince(i) = Fld(f, 5)
            mWhoIp(i) = Fld(f, 6)
            mWhoMs(i) = Fld(f, 7)
            mWhoLink(i) = Fld(f, 8)
            If mWhoLink(i) = "0" Then mWhoLink(i) = ""
            i = i + 1
        Next
    End If
    mWhoCount = n
    Set mWhoBuild = Nothing
    ' people who left today (only servers from 0.4.1 send these)
    n = 0
    If Not mGoneBuild Is Nothing Then n = mGoneBuild.Count
    If n > 0 Then
        ReDim mGoneNick(0 To n - 1)
        ReDim mGoneSeat(0 To n - 1)
        ReDim mGoneHow(0 To n - 1)
        ReDim mGoneAt(0 To n - 1)
        i = 0
        For Each v In mGoneBuild
            f = Split(CStr(v), "|")
            mGoneNick(i) = Fld(f, 1)
            mGoneSeat(i) = Fld(f, 2)
            If mGoneSeat(i) = "-" Then mGoneSeat(i) = ""
            mGoneHow(i) = Fld(f, 3)
            mGoneAt(i) = Fld(f, 4)
            i = i + 1
        Next
    End If
    mGoneCount = n
    Set mGoneBuild = Nothing
    n = mWhoCount
    For i = 1 To n - 1
        j = i
        Do While j > 0
            If WhoRank(j - 1) <= WhoRank(j) Then Exit Do
            SwapWho j - 1, j
            j = j - 1
        Loop
    Next
    DrawWhoGroup
    UpdateWhoScroll
    DrawWho
    DrawHeader
    DrawStatus
    BuildPeopleRows
    UpdatePeopleScroll
    If mPage = PG_WHO Then
        DrawWhoPage
        AskWhois
        RefreshWhoTools
    End If
    If Len(mWelcome) > 0 Then
        AddSys IC_OK, mWelcome, C_GREEN
        mWelcome = ""
        If Len(mWelcomeNote) > 0 Then AddSys IC_INFO, mWelcomeNote, C_BLUE
        mWelcomeNote = ""
    End If
End Sub

' ================================================================ the Who's Here page
' Two views, picked with the tabs at the top:
'   List      everyone, like Explorer's Details view, with the people who've left greyed out below
'   Seat map  the venue's tables with everyone in their seat (the host sets up a map per venue)
' On the right, a card about the person picked: seat, PC, ping and setup (from WHOIS).

' The same sizes for laying out and drawing
Private Sub WhoGeometry(w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long)
    w = picWhoPage.ScaleWidth
    h = picWhoPage.ScaleHeight
    rw = 230
    If w - 30 - rw < 300 Then rw = ((w - 30) * 2) \ 5
    lw = w - 30 - rw
    rx = 10 + lw + 10
    gH = h - 16 - 10 - 34 - 10
    If gH < 120 Then gH = 120
End Sub

Private Sub LayoutWho()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long, by As Long, ph As Long
    WhoGeometry w, h, lw, rx, rw, gH
    ph = 16 + gH - TAB_BOTTOM
    picPeople.Move 16, TAB_BOTTOM + 6, lw - 12, ph - 12
    vsbPeople.Move picPeople.ScaleWidth - 16, 0, 16, picPeople.ScaleHeight
    picMap.Move 16, TAB_BOTTOM + 6, lw - 12, ph - 12
    vsbMap.Move picMap.ScaleWidth - 16, 0, 16, picMap.ScaleHeight
    picPeople.Visible = (mWhoView = 0)
    picMap.Visible = (mWhoView = 1)
    by = 16 + gH - 8 - 25
    picCard.Move rx + 6, 28, rw - 12, by - 6 - 28
    cmdWhoSeat.Move rx + 8, by, 130, 25
    cmdWhoChat.Move rx + 8, by, 150, 25
    UpdatePeopleScroll
    DrawWhoPage
End Sub

Private Sub DrawWhoPage()
    Want DR_WHOPAGE
End Sub

Private Sub DrawWhoPageNow()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long, s As String, ic As Integer
    WhoGeometry w, h, lw, rx, rw, gH
    picWhoPage.Cls
    DrawWhoTabs lw, 16 + gH
    If Len(mPplSel) = 0 Then
        s = "About the User"
    ElseIf StrComp(mPplSel, mMyNick, vbTextCompare) = 0 Then
        s = "About You"
    Else
        s = "About " & mPplSel
    End If
    GroupBox picWhoPage, rx, 16, rw, gH, IC_USER, s
    TipBox picWhoPage, 10, h - 10 - 34, w - 20, 34
    If mWhoView = 1 And (Len(mMySeat) > 0 Or Not mSignedIn) Then
        ' the key to the map
        DrawMapKey picWhoPage, 18, h - 10 - 34 + 9
    Else
        If mSignedIn And Len(mMySeat) = 0 Then
            ic = IC_WARN
            If mWhoView = 1 Then
                s = "No seat selected. Click a free seat on the map to select it."
            Else
                s = "No seat selected. Click Change Seat on the toolbar to select one."
            End If
        Else
            ic = IC_INFO
            s = "Double-click a name to chat with that user. Click a column heading to sort."
        End If
        DrawIcon16 picWhoPage, ic, 18, h - 10 - 34 + 9
        TextAt picWhoPage, 40, h - 10 - 34 + 10, FitText(picWhoPage, s, w - 60), vbBlack
    End If
    If mWhoView = 1 Then
        DrawMapViewNow
    Else
        DrawPeopleNow
    End If
    DrawCardNow
    UpdateWhoButtons
End Sub

' ---- the tabs

Private Function TabText(ByVal i As Integer) As String
    If i = 0 Then
        TabText = "List"
    Else
        TabText = "Seat Map"
    End If
End Function

Private Function TabIcon(ByVal i As Integer) As Integer
    If i = 0 Then
        TabIcon = IC_PEOPLE
    Else
        TabIcon = IC_PLAN
    End If
End Function

' Where tab i starts and how wide it is
Private Sub TabSpan(ByVal i As Integer, x As Long, w As Long)
    Dim k As Integer
    x = 12
    For k = 0 To i
        picWhoPage.FontBold = True
        w = 6 + 16 + 5 + picWhoPage.TextWidth(TabText(k)) + 10
        picWhoPage.FontBold = False
        If k < i Then x = x + w + 1
    Next
End Sub

' A raised panel with the two tabs on top, like a Windows tab control
Private Sub DrawWhoTabs(ByVal lw As Long, ByVal panelBottom As Long)
    Dim pass As Integer, i As Integer, x As Long, w As Long, y As Long, h As Long
    Edge picWhoPage, 10, TAB_BOTTOM, lw, panelBottom - TAB_BOTTOM, vb3DHighlight, vb3DDKShadow
    Edge picWhoPage, 11, TAB_BOTTOM + 1, lw - 2, panelBottom - TAB_BOTTOM - 2, vbButtonFace, vb3DShadow
    ' the other tab first, then the one picked, so the picked one sits on top
    For pass = 0 To 1
        For i = 0 To 1
            If (i = mWhoView) = (pass = 1) Then
                TabSpan i, x, w
                If i = mWhoView Then
                    x = x - 2
                    w = w + 4
                    y = TAB_BOTTOM - 22
                    h = 23
                Else
                    y = TAB_BOTTOM - 20
                    h = 20
                End If
                Box picWhoPage, x, y, w, h, vbButtonFace
                Box picWhoPage, x, y + 2, 1, h - 2, vb3DHighlight
                Box picWhoPage, x + 1, y + 1, 1, 1, vb3DHighlight
                Box picWhoPage, x + 2, y, w - 4, 1, vb3DHighlight
                Box picWhoPage, x + w - 2, y + 1, 1, 1, vb3DDKShadow
                Box picWhoPage, x + w - 1, y + 2, 1, h - 2, vb3DDKShadow
                Box picWhoPage, x + w - 2, y + 2, 1, h - 2, vb3DShadow
                DrawIcon16 picWhoPage, TabIcon(i), x + 6, y + 3
                TextAt picWhoPage, x + 27, y + 5, TabText(i), vbButtonText, (i = mWhoView)
            End If
        Next
    Next
End Sub

Private Function TabAt(ByVal px As Single, ByVal py As Single) As Integer
    Dim i As Integer, x As Long, w As Long
    TabAt = -1
    If py < TAB_BOTTOM - 22 Or py >= TAB_BOTTOM Then Exit Function
    For i = 0 To 1
        TabSpan i, x, w
        If px >= x And px < x + w Then TabAt = i
    Next
End Function

Private Sub picWhoPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    If Button <> vbLeftButton Then Exit Sub
    i = TabAt(X, Y)
    If i >= 0 And i <> mWhoView Then SetWhoView i
End Sub

Private Sub SetWhoView(ByVal v As Integer)
    mWhoView = v
    SettingSet "who_view", CStr(v)
    picPeople.Visible = (v = 0)
    picMap.Visible = (v = 1)
    DrawWhoPage
    RefreshWhoTools
    On Error Resume Next
    If v = 1 Then
        picMap.SetFocus
    Else
        picPeople.SetFocus
    End If
End Sub

' ---- the list

' Puts everyone into rows: a heading, the people signed in, then a heading and the people who left
Private Sub BuildPeopleRows()
    Dim i As Long, idx() As Long, keys() As String
    ReDim mRowKind(0 To 1 + mWhoCount + mGoneCount)
    ReDim mRowIdx(0 To 1 + mWhoCount + mGoneCount)
    mRowCount = 0
    AddRow 0, 0
    If mWhoCount > 0 Then
        ReDim idx(0 To mWhoCount - 1)
        ReDim keys(0 To mWhoCount - 1)
        For i = 0 To mWhoCount - 1
            idx(i) = i
            keys(i) = PeopleKey(False, i)
        Next
        SortIdx idx, keys, mWhoCount
        For i = 0 To mWhoCount - 1
            AddRow 1, idx(i)
        Next
    End If
    If mGoneCount > 0 Then
        AddRow 0, 1
        ReDim idx(0 To mGoneCount - 1)
        ReDim keys(0 To mGoneCount - 1)
        For i = 0 To mGoneCount - 1
            idx(i) = i
            keys(i) = PeopleKey(True, i)
        Next
        SortIdx idx, keys, mGoneCount
        For i = 0 To mGoneCount - 1
            AddRow 2, idx(i)
        Next
    End If
End Sub

Private Sub AddRow(ByVal kind As Integer, ByVal i As Long)
    mRowKind(mRowCount) = kind
    mRowIdx(mRowCount) = i
    mRowCount = mRowCount + 1
End Sub

' What the list is sorted on, for one person
Private Function PeopleKey(ByVal gone As Boolean, ByVal i As Long) As String
    Dim nm As String, seat As String, st As String, p1 As Long, p2 As Long
    p1 = 99999
    p2 = 99999
    If gone Then
        nm = mGoneNick(i)
        seat = mGoneSeat(i)
        st = "4" & mGoneHow(i)
    Else
        nm = mWhoNick(i)
        seat = mWhoSeat(i)
        If mWhoAway(i) Then
            st = "3"
        ElseIf mWhoBusy(i) Then
            st = "2"
        Else
            st = "1"
        End If
        If Len(mWhoMs(i)) > 0 Then p1 = Val(mWhoMs(i))
        p2 = PeerValue(mWhoIp(i))
        If p2 < 0 Then p2 = 99999
    End If
    If Len(seat) = 0 Then seat = "~~~~"          ' no seat goes last
    Select Case mPplSort
    Case 1: PeopleKey = seat & "|" & LCase$(nm)
    Case 2: PeopleKey = st & "|" & LCase$(nm)
    Case 3: PeopleKey = Format$(p1, "00000") & "|" & LCase$(nm)
    Case 4: PeopleKey = Format$(p2, "00000") & "|" & LCase$(nm)
    Case Else: PeopleKey = LCase$(nm)
    End Select
End Function

Private Sub SortIdx(idx() As Long, keys() As String, ByVal n As Long)
    Dim i As Long, j As Long, ti As Long, tk As String, swapIt As Boolean
    For i = 1 To n - 1
        j = i
        Do While j > 0
            If mPplDesc Then
                swapIt = (keys(j - 1) < keys(j))
            Else
                swapIt = (keys(j - 1) > keys(j))
            End If
            If Not swapIt Then Exit Do
            ti = idx(j - 1)
            idx(j - 1) = idx(j)
            idx(j) = ti
            tk = keys(j - 1)
            keys(j - 1) = keys(j)
            keys(j) = tk
            j = j - 1
        Loop
    Next
End Sub

Private Function RowNick(ByVal r As Long) As String
    If r < 0 Or r >= mRowCount Then Exit Function
    Select Case mRowKind(r)
    Case 1: RowNick = mWhoNick(mRowIdx(r))
    Case 2: RowNick = mGoneNick(mRowIdx(r))
    End Select
End Function

' Everything one row shows
Private Sub RowInfo(ByVal r As Long, nm As String, seat As String, st As String, p1 As String, p2 As String, ic As Integer, col As Long, warn As Boolean)
    Dim i As Long
    i = mRowIdx(r)
    warn = False
    p1 = ""
    p2 = ""
    If mRowKind(r) = 1 Then
        nm = mWhoNick(i)
        seat = mWhoSeat(i)
        p1 = MsText(mWhoMs(i))
        p2 = PeerText(i)
        warn = (Len(HandReason(mWhoLink(i), mWhoMs(i))) > 0)
        If mWhoAway(i) Then
            st = "Away"
            ic = IC_MOON
            col = C_GREY
        ElseIf mWhoBusy(i) Then
            st = "Busy"
            ic = IC_BAD
            col = vbBlack
        ElseIf mWhoAdmin(i) Then
            st = "Online (host)"
            ic = IC_ADMIN
            col = C_PURPLE
        Else
            st = "Online"
            ic = IC_USER
            col = vbBlack
        End If
    Else
        nm = mGoneNick(i)
        seat = mGoneSeat(i)
        st = GoneSentence(mGoneHow(i), mGoneAt(i))
        ic = GoneIcon(mGoneHow(i))
        col = C_GREY
    End If
End Sub

Private Function GoneWord(ByVal how As String) As String
    Select Case how
    Case "home": GoneWord = "Signed out"
    Case "quit": GoneWord = "Exited"
    Case "dropped": GoneWord = "Connection lost"
    Case "update": GoneWord = "Updating"
    Case "break": GoneWord = "Away"
    Case Else: GoneWord = "Disconnected"
    End Select
End Function

' "Signed out at 9:15 PM", "Away since 3:10 PM"
Private Function GoneSentence(ByVal how As String, ByVal leftAt As String) As String
    GoneSentence = GoneWord(how)
    If Len(leftAt) > 0 Then
        If how = "break" Then
            GoneSentence = GoneSentence & " since " & AmPm(leftAt)
        Else
            GoneSentence = GoneSentence & " at " & AmPm(leftAt)
        End If
    End If
End Function

Private Function GoneIcon(ByVal how As String) As Integer
    Select Case how
    Case "quit": GoneIcon = IC_PCOFF
    Case "dropped": GoneIcon = IC_SIG0
    Case "update": GoneIcon = IC_DOWNLOAD
    Case "break": GoneIcon = IC_MOON
    Case Else: GoneIcon = IC_DOOR
    End Select
End Function

Private Sub CalcCols(ByVal w As Long)
    Dim nameW As Long
    nameW = w - COL_SEAT_W - COL_STATUS_W - COL_PING_W - COL_YOU_W
    If nameW < 90 Then nameW = 90
    mColX(0) = 0
    mColX(1) = nameW
    mColX(2) = mColX(1) + COL_SEAT_W
    mColX(3) = mColX(2) + COL_STATUS_W
    mColX(4) = mColX(3) + COL_PING_W
    mColX(5) = w
    If mColX(5) < mColX(4) + 30 Then mColX(5) = mColX(4) + 30
End Sub

Private Function ColName(ByVal c As Integer) As String
    Select Case c
    Case 0: ColName = "Name"
    Case 1: ColName = "Seat"
    Case 2: ColName = "Status"
    Case 3: ColName = "Server Ping"
    Case Else: ColName = "Your Ping"
    End Select
End Function

Private Function PeopleVisible() As Long
    PeopleVisible = (picPeople.ScaleHeight - HDR_H - 1) \ PPL_ROW
    If PeopleVisible < 1 Then PeopleVisible = 1
End Function

Private Sub UpdatePeopleScroll()
    Dim vis As Long, mx As Long
    vis = PeopleVisible()
    mx = mRowCount - vis
    If mx < 0 Then mx = 0
    mIgnoreScroll = True
    If vsbPeople.Value > mx Then vsbPeople.Value = mx
    vsbPeople.Max = mx
    If vis > 1 Then
        vsbPeople.LargeChange = vis - 1
    Else
        vsbPeople.LargeChange = 1
    End If
    vsbPeople.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

Private Sub DrawPeople()
    Want DR_WHOBODY
End Sub

Private Sub DrawPeopleNow()
    Dim w As Long, c As Integer, r As Long, y As Long, s As String, tw As Long, nw As Long
    Dim nm As String, seat As String, st As String, p1 As String, p2 As String, ic As Integer, col As Long, warn As Boolean
    picPeople.Cls
    w = picPeople.ScaleWidth - 16
    CalcCols w
    For c = 0 To 4
        DrawColHead c
    Next
    For r = vsbPeople.Value To mRowCount - 1
        y = HDR_H + 1 + (r - vsbPeople.Value) * PPL_ROW
        If y > picPeople.ScaleHeight Then Exit For
        If mRowKind(r) = 0 Then
            ' a group heading, with a line after it
            If mRowIdx(r) = 0 Then
                s = "Signed In (" & mWhoCount & ")"
            Else
                s = "Signed Out (" & mGoneCount & ")"
            End If
            TextAt picPeople, 6, y + 3, s, C_TITLE, True
            picPeople.FontBold = True
            tw = picPeople.TextWidth(s)
            picPeople.FontBold = False
            Box picPeople, 6 + tw + 6, y + 9, w - (6 + tw + 6) - 6, 1, vb3DShadow
        Else
            RowInfo r, nm, seat, st, p1, p2, ic, col, warn
            If StrComp(nm, mPplSel, vbTextCompare) = 0 Then
                Box picPeople, 1, y, w - 2, PPL_ROW, C_SEL
                col = vbWhite
            End If
            DrawIcon16 picPeople, ic, mColX(0) + 8, y + 1
            s = nm
            If mRowKind(r) = 1 And StrComp(nm, mMyNick, vbTextCompare) = 0 Then s = s & " (you)"
            nw = mColX(1) - mColX(0) - 32
            If warn Then
                DrawIcon16 picPeople, IC_WARN, mColX(1) - 20, y + 1
                nw = nw - 18
            End If
            TextAt picPeople, mColX(0) + 28, y + 3, FitText(picPeople, s, nw), col
            TextAt picPeople, mColX(1) + 6, y + 3, FitText(picPeople, seat, mColX(2) - mColX(1) - 8), col
            If mRowKind(r) = 2 Then
                ' people who left: what happened and when, across the status and ping columns
                TextAt picPeople, mColX(2) + 6, y + 3, FitText(picPeople, st, mColX(5) - mColX(2) - 8), col
            Else
                TextAt picPeople, mColX(2) + 6, y + 3, FitText(picPeople, st, mColX(3) - mColX(2) - 8), col
                TextAt picPeople, mColX(3) + 6, y + 3, FitText(picPeople, p1, mColX(4) - mColX(3) - 8), col
                TextAt picPeople, mColX(4) + 6, y + 3, FitText(picPeople, p2, mColX(5) - mColX(4) - 8), col
            End If
        End If
    Next
End Sub

' One column heading, a raised button like Explorer's. Pushed in while you press it.
Private Sub DrawColHead(ByVal c As Integer)
    Dim x As Long, w As Long, s As String, d As Long, tx As Long
    x = mColX(c)
    w = mColX(c + 1) - x
    Box picPeople, x, 0, w, HDR_H, vbButtonFace
    If mHdrDown = c Then
        Edge picPeople, x, 0, w, HDR_H, vb3DShadow, vb3DShadow
        d = 1
    Else
        Edge picPeople, x, 0, w, HDR_H, vb3DHighlight, vb3DDKShadow
        Box picPeople, x + 1, HDR_H - 2, w - 2, 1, vb3DShadow
        Box picPeople, x + w - 2, 1, 1, HDR_H - 2, vb3DShadow
    End If
    s = FitText(picPeople, ColName(c), w - 12)
    TextAt picPeople, x + 6 + d, 3 + d, s, vbButtonText
    ' a small triangle on the column the list is sorted by
    If c = mPplSort Then
        tx = x + 6 + picPeople.TextWidth(s) + 7 + d
        If tx + 7 < x + w - 3 Then DrawSortArrow tx, 7 + d
    End If
End Sub

Private Sub DrawSortArrow(ByVal x As Long, ByVal y As Long)
    Dim i As Long
    For i = 0 To 3
        If mPplDesc Then
            Box picPeople, x + i, y + i, 7 - 2 * i, 1, vb3DShadow
        Else
            Box picPeople, x + 3 - i, y + i, 1 + 2 * i, 1, vb3DShadow
        End If
    Next
End Sub

Private Function ColAt(ByVal px As Single) As Integer
    Dim c As Integer
    ColAt = -1
    For c = 0 To 4
        If px >= mColX(c) And px < mColX(c + 1) Then ColAt = c
    Next
End Function

Private Function RowAt(ByVal py As Single) As Long
    Dim r As Long
    RowAt = -1
    If py < HDR_H + 1 Then Exit Function
    r = vsbPeople.Value + Int((py - HDR_H - 1) / PPL_ROW)
    If r < mRowCount Then RowAt = r
End Function

Private Sub picPeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Long
    On Error Resume Next
    picPeople.SetFocus
    On Error GoTo 0
    mPplDownRow = -1
    If Y < HDR_H Then
        ' a column heading: sorted when the button comes back up
        If Button = vbLeftButton Then
            mHdrDown = ColAt(X)
            DrawPeople
        End If
        Exit Sub
    End If
    r = RowAt(Y)
    If r < 0 Then Exit Sub
    If mRowKind(r) = 0 Then Exit Sub
    mPplDownRow = r
    SelectPerson RowNick(r)
End Sub

Private Sub picPeople_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Integer
    If mHdrDown < 0 Then Exit Sub
    c = -1
    If Y >= 0 And Y < HDR_H Then c = ColAt(X)
    If c >= 0 And c = mHdrDown Then
        If mPplSort = c Then
            mPplDesc = Not mPplDesc
        Else
            mPplSort = c
            mPplDesc = False
        End If
        SettingSet "who_sort", CStr(mPplSort)
        If mPplDesc Then
            SettingSet "who_desc", "1"
        Else
            SettingSet "who_desc", "0"
        End If
        BuildPeopleRows
    End If
    mHdrDown = -1
    DrawPeople
End Sub

Private Sub picPeople_DblClick()
    If mPplDownRow < 0 Then Exit Sub
    If CanChatTo(mPplSel) Then ChatTo mPplSel
End Sub

' Up and down move through the list, Enter chats to them
Private Sub picPeople_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long, cur As Long, stp As Long
    cur = -1
    For r = 0 To mRowCount - 1
        If mRowKind(r) <> 0 Then
            If StrComp(RowNick(r), mPplSel, vbTextCompare) = 0 Then cur = r
        End If
    Next
    Select Case KeyCode
    Case vbKeyUp
        stp = -1
    Case vbKeyDown
        stp = 1
    Case vbKeyReturn
        If CanChatTo(mPplSel) Then ChatTo mPplSel
        Exit Sub
    Case Else
        Exit Sub
    End Select
    r = cur + stp
    Do While r >= 0 And r < mRowCount
        If mRowKind(r) <> 0 Then
            If r < vsbPeople.Value Then vsbPeople.Value = r
            If r >= vsbPeople.Value + PeopleVisible() Then vsbPeople.Value = r - PeopleVisible() + 1
            SelectPerson RowNick(r)
            Exit Sub
        End If
        r = r + stp
    Loop
End Sub

Private Sub vsbPeople_Change()
    If Not mIgnoreScroll Then DrawPeople
End Sub

Private Sub vsbPeople_Scroll()
    DrawPeople
End Sub

' ---- the seat map

Private Function ServerHasMaps() As Boolean
    ServerHasMaps = (VerNum(mServerVer) >= 402)
End Function

Private Sub DrawMapView()
    Want DR_WHOBODY
End Sub

Private Sub DrawMapViewNow()
    Dim w As Long, y As Long, s As String, extra As String
    picMap.Cls
    w = picMap.ScaleWidth - 16
    If mMap.Count = 0 Then
        mLay.SeatCount = 0
        mLay.BanCount = 0
        mMapH = 0
        UpdateMapScroll
        DrawIcon32 picMap, IC_PLAN, (w - 32) \ 2, 30
        If ServerHasMaps() Or Len(mServerVer) = 0 Then
            s = "No seat map has been set up for this venue."
            TextAt picMap, (w - picMap.TextWidth(s)) \ 2, 72, s, C_SUB
            If mIsAdmin Then
                s = "Click Seat Map on the toolbar to set one up."
            Else
                s = "The host can set one up."
            End If
        Else
            s = "Seat maps require a newer Palms Server."
            TextAt picMap, (w - picMap.TextWidth(s)) \ 2, 72, s, C_SUB
            s = "Update palms_server.py to version 0.4.2 or later."
        End If
        TextAt picMap, (w - picMap.TextWidth(s)) \ 2, 90, FitText(picMap, s, w - 8), C_GREY
        Exit Sub
    End If
    LayoutMap mMap, w, False, mLay
    FillSeats
    extra = OffMapText()
    mMapH = mLay.Height
    If Len(extra) > 0 Then mMapH = mMapH + 24
    UpdateMapScroll
    DrawMap picMap, mLay, vsbMap.Value
    If Len(extra) > 0 Then
        y = mLay.Height - vsbMap.Value
        DrawIcon16 picMap, IC_INFO, 8, y
        TextAt picMap, 28, y + 2, FitText(picMap, extra, w - 36), C_SUB
    End If
End Sub

Private Sub UpdateMapScroll()
    Dim mx As Long
    mx = mMapH - picMap.ScaleHeight
    If mx < 0 Then mx = 0
    If mx > 30000 Then mx = 30000
    mIgnoreScroll = True
    If vsbMap.Value > mx Then vsbMap.Value = mx
    vsbMap.Max = mx
    vsbMap.SmallChange = 20
    If picMap.ScaleHeight > 40 Then
        vsbMap.LargeChange = picMap.ScaleHeight - 20
    Else
        vsbMap.LargeChange = 20
    End If
    vsbMap.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

' Who's in each seat: people signed in, then seats kept for people who stepped out
Private Sub FillSeats()
    Dim k As Long, i As Long, id As String
    For k = 0 To mLay.SeatCount - 1
        id = mLay.SeatId(k)
        mLay.SeatState(k) = MS_FREE
        mLay.SeatName(k) = ""
        mLay.SeatWarn(k) = False
        mLay.SeatSel(k) = False
        For i = 0 To mGoneCount - 1
            If mGoneSeat(i) = id And mGoneHow(i) <> "home" Then
                mLay.SeatState(k) = MS_HELD
                mLay.SeatName(k) = mGoneNick(i)
                mLay.SeatSel(k) = (StrComp(mGoneNick(i), mPplSel, vbTextCompare) = 0)
            End If
        Next
        For i = 0 To mWhoCount - 1
            If mWhoSeat(i) = id Then
                mLay.SeatName(k) = mWhoNick(i)
                If StrComp(mWhoNick(i), mMyNick, vbTextCompare) = 0 Then
                    mLay.SeatState(k) = MS_YOU
                ElseIf mWhoAway(i) Then
                    mLay.SeatState(k) = MS_AWAY
                ElseIf mWhoAdmin(i) Then
                    mLay.SeatState(k) = MS_HOST
                Else
                    mLay.SeatState(k) = MS_TAKEN
                End If
                mLay.SeatWarn(k) = (Len(HandReason(mWhoLink(i), mWhoMs(i))) > 0)
                mLay.SeatSel(k) = (StrComp(mWhoNick(i), mPplSel, vbTextCompare) = 0)
            End If
        Next
    Next
End Sub

' "Not on the map: b0b (no seat), zed (Z99)"
Private Function OffMapText() As String
    Dim i As Long, s As String
    For i = 0 To mWhoCount - 1
        If Len(mWhoSeat(i)) = 0 Then
            If Len(s) > 0 Then s = s & ", "
            s = s & mWhoNick(i) & " (no seat)"
        ElseIf SeatIndex(mLay, mWhoSeat(i)) < 0 Then
            If Len(s) > 0 Then s = s & ", "
            s = s & mWhoNick(i) & " (" & mWhoSeat(i) & ")"
        End If
    Next
    If Len(s) > 0 Then OffMapText = "Not on the map: " & s
End Function

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SeatAt(mLay, X, Y + vsbMap.Value) >= 0 Then HandCursor
End Sub

' Click someone to see their card. Click a free seat to sit there.
Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k As Long
    On Error Resume Next
    picMap.SetFocus
    On Error GoTo 0
    If Button <> vbLeftButton Then Exit Sub
    k = SeatAt(mLay, X, Y + vsbMap.Value)
    mMapDownSeat = k
    If k < 0 Then Exit Sub
    HandCursor
    If mLay.SeatState(k) <> MS_FREE Then SelectPerson mLay.SeatName(k)
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k As Long
    If Button <> vbLeftButton Then Exit Sub
    k = SeatAt(mLay, X, Y + vsbMap.Value)
    If k < 0 Or k <> mMapDownSeat Then Exit Sub
    If mLay.SeatState(k) = MS_FREE Then SitAt mLay.SeatId(k)
End Sub

Private Sub picMap_DblClick()
    Dim k As Long
    k = mMapDownSeat
    If k < 0 Or k >= mLay.SeatCount Then Exit Sub
    If mLay.SeatState(k) = MS_FREE Then Exit Sub
    If CanChatTo(mLay.SeatName(k)) Then ChatTo mLay.SeatName(k)
End Sub

Private Sub SitAt(ByVal id As String)
    Dim q As String
    If Not ServerIsNew() Or Not mSignedIn Then Exit Sub
    If id = mMySeat Then Exit Sub
    If Len(mMySeat) > 0 Then
        q = "Move from seat " & mMySeat & " to seat " & id & "?"
    Else
        q = "Select seat " & id & "?"
    End If
    If MsgBox(q, vbQuestion + vbYesNo, "Seat Map") = vbYes Then SendLine "SEAT|" & id
End Sub

Private Sub vsbMap_Change()
    If Not mIgnoreScroll Then DrawMapView
End Sub

Private Sub vsbMap_Scroll()
    DrawMapView
End Sub

Private Sub ShowMapEditor()
    If Not (mSignedIn And mIsAdmin) Then Exit Sub
    If Not ServerHasMaps() Then
        MsgBox "Seat maps require Palms Server version 0.4.2 or later.", vbInformation, "Seat Map"
        Exit Sub
    End If
    frmMapEdit.Show vbModal, Me
End Sub

' ---- picking someone

Private Sub SelectPerson(ByVal nick As String)
    If StrComp(nick, mPplSel, vbTextCompare) <> 0 Then
        mPplSel = nick
        ReDim mCard(0 To 0)
        mCardNick = ""
    End If
    DrawWhoPage
    RefreshWhoTools
    AskWhois
End Sub

' Ask the server about the person picked (servers from 0.4.1 know how)
Private Sub AskWhois()
    If Len(mPplSel) > 0 And ServerIsNew() Then SendLine "WHOIS|" & mPplSel
End Sub

Private Function PersonIsHere(ByVal nick As String) As Boolean
    Dim i As Long
    If Len(nick) = 0 Then Exit Function
    For i = 0 To mWhoCount - 1
        If StrComp(mWhoNick(i), nick, vbTextCompare) = 0 Then PersonIsHere = True
    Next
End Function

Private Function WhoIndex(ByVal nick As String) As Long
    Dim i As Long
    WhoIndex = -1
    If Len(nick) = 0 Then Exit Function
    For i = 0 To mWhoCount - 1
        If StrComp(mWhoNick(i), nick, vbTextCompare) = 0 Then WhoIndex = i
    Next
End Function

Private Function CanChatTo(ByVal nick As String) As Boolean
    CanChatTo = PersonIsHere(nick) And StrComp(nick, mMyNick, vbTextCompare) <> 0
End Function

' Go to Chat with their name typed in, ready for you to say something
Private Sub ChatTo(ByVal nick As String)
    Dim s As String
    s = Trim$(txtSay.Text)
    If Len(s) > 0 Then s = " " & s
    GoPage PG_CHAT
    txtSay.Text = nick & ":" & s & " "
    txtSay.SelStart = Len(txtSay.Text)
    On Error Resume Next
    txtSay.SetFocus
End Sub

Private Sub UpdateWhoButtons()
    Dim isMe As Boolean
    isMe = (Len(mPplSel) > 0 And StrComp(mPplSel, mMyNick, vbTextCompare) = 0)
    cmdWhoSeat.Visible = isMe And mSignedIn
    cmdWhoChat.Visible = CanChatTo(mPplSel)
    If cmdWhoChat.Visible Then cmdWhoChat.Caption = "&Chat with " & mPplSel
End Sub

Private Sub RefreshWhoTools()
    If mPage = PG_WHO Then
        BuildTools
        DrawTools
    End If
End Sub

Private Sub cmdWhoChat_Click()
    If CanChatTo(mPplSel) Then ChatTo mPplSel
End Sub

Private Sub cmdWhoSeat_Click()
    ShowSeatDialog
End Sub

' ---- ping: from the server (it measures everyone) and from this PC (we measure, see PingPeers)

' "3 ms", "<1 ms", or "-" when we don't know
Private Function MsWords(ByVal ms As Long) As String
    If ms < 1 Then
        MsWords = "<1 ms"
    Else
        MsWords = ms & " ms"
    End If
End Function

Private Function MsText(ByVal s As String) As String
    If Len(s) = 0 Then
        MsText = "-"
    Else
        MsText = MsWords(Val(s))
    End If
End Function

' Why someone might need a hand, or "" if they look fine
Private Function HandReason(ByVal link As String, ByVal ms As String) As String
    If Val(link) > 0 And Val(link) < SLOW_LINK_MB Then
        HandReason = "Link speed is " & Val(link) & " Mbps. Check the network cable."
    ElseIf Len(ms) > 0 And Val(ms) >= SLOW_PING_MS Then
        HandReason = "High ping (" & Val(ms) & " ms). Check the network cable."
    End If
End Function

Private Function PeerFind(ByVal ip As String) As Long
    Dim k As Long
    PeerFind = -1
    For k = 0 To mPeerCount - 1
        If mPeerIp(k) = ip Then
            PeerFind = k
            Exit Function
        End If
    Next
End Function

Private Function PeerSlot(ByVal ip As String) As Long
    PeerSlot = PeerFind(ip)
    If PeerSlot >= 0 Then Exit Function
    If mPeerCount > MAX_PEERS - 1 Then Exit Function
    PeerSlot = mPeerCount
    mPeerIp(PeerSlot) = ip
    mPeerMs(PeerSlot) = 0
    mPeerAt(PeerSlot) = 0
    mPeerSent(PeerSlot) = 0
    mPeerCount = mPeerCount + 1
End Function

' Our last ping to that PC in ms, or -1 if it hasn't answered lately
Private Function PeerValue(ByVal ip As String) As Long
    Dim k As Long
    PeerValue = -1
    k = PeerFind(ip)
    If k < 0 Then Exit Function
    If mPeerAt(k) = 0 Then Exit Function
    If GetTickCount() - mPeerAt(k) > 20000 Then Exit Function
    PeerValue = mPeerMs(k)
End Function

' "Your ping" for someone signed in
Private Function PeerText(ByVal i As Long) As String
    Dim ip As String, v As Long, k As Long
    PeerText = "-"
    If StrComp(mWhoNick(i), mMyNick, vbTextCompare) = 0 Then Exit Function
    ip = mWhoIp(i)
    If Len(ip) = 0 Or Left$(ip, 4) = "127." Then Exit Function
    v = PeerValue(ip)
    If v >= 0 Then
        PeerText = MsWords(v)
        Exit Function
    End If
    k = PeerFind(ip)
    PeerText = "..."
    If k >= 0 Then
        If mPeerSent(k) <> 0 And GetTickCount() - mPeerSent(k) > 3000 Then PeerText = "Timed out"
    End If
End Function

' Sends a tiny "ping" to every other LAN Center on UDP port 4099. They answer
' straight away (see sckFind_DataArrival) and the time it took is our ping to them.
Private Sub PingPeers()
    Dim i As Long, ip As String, k As Long, myIp As String
    On Error Resume Next
    myIp = sckFind.LocalIP
    For i = 0 To mWhoCount - 1
        ip = mWhoIp(i)
        If Len(ip) > 0 And Left$(ip, 4) <> "127." And ip <> myIp And StrComp(mWhoNick(i), mMyNick, vbTextCompare) <> 0 Then
            k = PeerSlot(ip)
            If k >= 0 Then
                mPeerSent(k) = GetTickCount()
                SendUdp ip, "PPING|" & mPeerSent(k)
            End If
        End If
    Next
End Sub

Private Sub PeerAnswered(ByVal ip As String, ByVal sent As Double)
    Dim k As Long, ms As Double
    ms = GetTickCount() - sent
    If ms < 0 Or ms > 10000 Then Exit Sub
    k = PeerSlot(ip)
    If k < 0 Then Exit Sub
    mPeerMs(k) = CLng(ms)
    mPeerAt(k) = GetTickCount()
    mPeerDirty = True
End Sub

Private Sub SendUdp(ByVal ip As String, ByVal s As String)
    On Error Resume Next
    If sckFind.State <> sckOpen Then Exit Sub
    sckFind.RemoteHost = ip
    sckFind.RemotePort = BEACON_PORT
    sckFind.SendData s
End Sub

' ---- the card about one person

Private Function CardVal(ByVal i As Long) As String
    If StrComp(mCardNick, mPplSel, vbTextCompare) <> 0 Then Exit Function
    If i <= UBound(mCard) Then CardVal = mCard(i)
End Function

' Name, seat, a line about where they are, an icon and a colour for that line
Private Sub CardBasics(nm As String, seat As String, st As String, ic As Integer, col As Long)
    Dim i As Long, how As String, leftAt As String, since As String, admin As Boolean, found As Boolean
    nm = mPplSel
    i = WhoIndex(mPplSel)
    If i >= 0 Then
        nm = mWhoNick(i)
        seat = mWhoSeat(i)
        since = mWhoSince(i)
        admin = mWhoAdmin(i)
        If mWhoAway(i) Then
            how = "away"
        ElseIf mWhoBusy(i) Then
            how = "busy"
        Else
            how = "here"
        End If
        found = True
    End If
    If Not found Then
        For i = 0 To mGoneCount - 1
            If StrComp(mGoneNick(i), mPplSel, vbTextCompare) = 0 Then
                nm = mGoneNick(i)
                seat = mGoneSeat(i)
                how = mGoneHow(i)
                leftAt = mGoneAt(i)
                found = True
            End If
        Next
    End If
    If Not found And Len(CardVal(1)) > 0 Then
        ' only the server's answer knows about them
        nm = CardVal(1)
        seat = CardVal(2)
        If seat = "-" Then seat = ""
        how = CardVal(3)
        leftAt = CardVal(6)
    End If
    If Len(since) = 0 Then since = CardVal(5)
    Select Case how
    Case "here"
        st = "Online"
        If Len(since) > 0 Then st = "Online since " & AmPm(since)
        If admin Then
            st = st & ", host"
            ic = IC_ADMIN
            col = C_PURPLE
        Else
            ic = IC_USER
            col = C_GREEN
        End If
    Case "away"
        st = "Away"
        ic = IC_MOON
        col = C_GREY
    Case "busy"
        st = "Busy"
        If admin Then st = st & ", host"
        ic = IC_BAD
        col = C_RED
    Case "home", "quit", "dropped", "break", "update"
        st = GoneSentence(how, leftAt)
        ic = GoneIcon(how)
        col = C_GREY
    Case Else
        st = "Not signed in"
        ic = IC_USER
        col = C_GREY
    End Select
End Sub

Private Sub DrawCard()
    Want DR_CARD
End Sub

Private Sub DrawCardNow()
    Dim w As Long, y As Long, s As String, ic As Integer, col As Long, i As Long, isMe As Boolean
    Dim nm As String, seat As String, st As String, why As String, link As String, ms As String
    picCard.Cls
    w = picCard.ScaleWidth
    If Len(mPplSel) = 0 Then
        DrawIcon32 picCard, IC_PEOPLE, (w - 32) \ 2, 30
        s = "Select a user in the list"
        TextAt picCard, (w - picCard.TextWidth(s)) \ 2, 72, s, C_SUB
        s = "to view details."
        TextAt picCard, (w - picCard.TextWidth(s)) \ 2, 88, s, C_SUB
        Exit Sub
    End If
    isMe = (StrComp(mPplSel, mMyNick, vbTextCompare) = 0)
    CardBasics nm, seat, st, ic, col
    DrawIcon32 picCard, ic, 4, 4
    picCard.FontSize = 12
    TextAt picCard, 44, 2, FitText(picCard, nm, w - 48), C_TITLE, True
    picCard.FontSize = 8
    TextAt picCard, 45, 24, FitText(picCard, st, w - 48), col
    y = 44
    Box picCard, 2, y, w - 4, 1, vb3DShadow
    Box picCard, 2, y + 1, w - 4, 1, vb3DHighlight
    y = y + 7
    ' the server's ping and the network card speed come from the list, or from WHOIS
    i = WhoIndex(mPplSel)
    If i >= 0 Then
        link = mWhoLink(i)
        ms = mWhoMs(i)
    Else
        link = CardVal(17)
        ms = CardVal(18)
    End If
    If i >= 0 Then why = HandReason(link, ms)
    If Len(why) > 0 Then
        DrawIcon16 picCard, IC_WARN, 4, y
        TextAt picCard, 24, y + 2, FitText(picCard, why, w - 28), C_RED, True
        y = y + 19
    End If
    If Len(seat) > 0 Then
        CardRow y, "Seat", seat, vbBlack, True
    Else
        CardRow y, "Seat", "None", C_GREY, False
    End If
    If CardWaiting() Then
        CardRow y, "PC", "...", C_GREY, False
    Else
        s = CardVal(7)
        If Len(CardVal(8)) > 0 Then
            If Len(s) > 0 Then
                s = s & " (" & CardVal(8) & ")"
            Else
                s = CardVal(8)
            End If
        End If
        If Len(s) = 0 Then s = "-"
        CardRow y, "PC", s, vbBlack, False
    End If
    If i >= 0 Then
        CardRow y, "Server ping", MsText(ms), vbBlack, False
        If isMe Then
            CardRow y, "Your ping", "-", C_GREY, False
        Else
            CardRow y, "Your ping", PeerText(i), vbBlack, False
        End If
    End If
    CardField y, "LAN Center", 9, "-"
    y = y + 4
    DrawIcon16 picCard, IC_PC, 4, y
    If isMe Then
        s = "System Information"
    Else
        s = "System Information"
    End If
    TextAt picCard, 24, y + 2, s, vbBlack, True
    y = y + 19
    CardField y, "Processor", 10, "Not available"
    CardField y, "Memory", 11, "Not available"
    CardField y, "Graphics", 12, "Not available"
    CardField y, "Sound", 15, "Not available"
    If CardWaiting() Or Len(CardVal(16)) = 0 Then
        CardField y, "Network", 16, "Not available"
    Else
        s = CardVal(16)
        If Val(CardVal(17)) > 0 Then s = Val(CardVal(17)) & " Mb, " & s
        If Val(CardVal(17)) > 0 And Val(CardVal(17)) < SLOW_LINK_MB Then
            CardRow y, "Network", s, C_RED, False
        Else
            CardRow y, "Network", s, vbBlack, False
        End If
    End If
    CardField y, "Windows", 13, "Not available"
    CardField y, "Screen", 14, "Not available"
End Sub

' True until the server's answer about the person picked has arrived
Private Function CardWaiting() As Boolean
    CardWaiting = (StrComp(mCardNick, mPplSel, vbTextCompare) <> 0)
End Function

' One line of the card: a label, then a value from the server's answer
Private Sub CardField(y As Long, ByVal lbl As String, ByVal i As Long, ByVal emptyText As String)
    Dim v As String
    If CardWaiting() Then
        CardRow y, lbl, "...", C_GREY, False
        Exit Sub
    End If
    v = CardVal(i)
    If Len(v) = 0 Then
        CardRow y, lbl, emptyText, C_GREY, False
    Else
        CardRow y, lbl, v, vbBlack, False
    End If
End Sub

Private Sub CardRow(y As Long, ByVal lbl As String, ByVal txt As String, ByVal col As Long, ByVal bold As Boolean)
    Dim w As Long
    w = picCard.ScaleWidth
    TextAt picCard, 6, y, lbl, C_SUB
    picCard.FontBold = bold
    txt = FitText(picCard, txt, w - 82)
    picCard.FontBold = False
    TextAt picCard, 78, y, txt, col, bold
    y = y + 14
End Sub

' ================================================================ the Game Library page
' The games themselves (the list from the pak server, what's installed, installing) are
' looked after by modLibrary. This part draws the page and handles the clicks.

' Servers from 0.4.3 pass on the pak server's list of games
Private Function ServerHasPaks() As Boolean
    ServerHasPaks = (VerNum(mServerVer) >= 403)
End Function

Private Sub LibGeometry(w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long)
    w = picLibPage.ScaleWidth
    h = picLibPage.ScaleHeight
    rw = 236
    If w - 30 - rw < 300 Then rw = ((w - 30) * 2) \ 5
    lw = w - 30 - rw
    rx = 10 + lw + 10
    gH = h - 16 - 10 - 34 - 10
    If gH < 160 Then gH = 160
End Sub

Private Sub LayoutLibrary()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long, by As Long, bw As Long, ch As Long
    LibGeometry w, h, lw, rx, rw, gH
    picLibList.Move 16, 30, lw - 12, gH - 20
    vsbLib.Move picLibList.ScaleWidth - 16, 0, 16, picLibList.ScaleHeight
    ' the card, with the big button along the bottom and two smaller ones above it
    by = 16 + gH - 8 - 25
    bw = (rw - 16 - 6) \ 2
    cmdLibGo.Move rx + 8, by, rw - 16, 25
    cmdLibAlt.Move rx + 8, by - 31, bw, 25
    cmdLibDel.Move rx + 8 + bw + 6, by - 31, rw - 16 - bw - 6, 25
    ch = by - 31 - 6 - 28
    If ch < 40 Then ch = 40
    picLibCard.Move rx + 6, 28, rw - 12, ch
    UpdateLibScroll
    DrawLibPage
End Sub

' Something about the games changed: sort the list again and redraw it all
Private Sub LibRefresh()
    LibRebuild
    gLibDirty = False
    BuildLibRows
    UpdateLibScroll
    DrawLibPage
    If mPage = PG_LIBRARY Or mPage = PG_WHATSON Then
        BuildTools
        DrawTools
        DrawHeader
    End If
    LibNavBadge
    If LibWorking() Then LibWake
End Sub

Private Sub DrawLibPage()
    Want DR_LIBPAGE
End Sub

Private Sub DrawLibPageNow()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long, s As String, ic As Integer, free As Double
    LibGeometry w, h, lw, rx, rw, gH
    picLibPage.Cls
    GroupBox picLibPage, 10, 16, lw, gH, IC_DISC, "Games (" & gShowCount & ")"
    If LibShown(mLibSel) Then
        s = "About " & LibName(mLibSel)
    Else
        s = "About the Game"
    End If
    GroupBox picLibPage, rx, 16, rw, gH, IC_INFO, FitText(picLibPage, s, rw - 40)
    ' a tip along the bottom: problems first, otherwise where games go
    TipBox picLibPage, 10, h - 10 - 34, w - 20, 34
    s = ToolsMissing()
    If Len(s) > 0 Then
        ic = IC_WARN
        s = "This PC is missing " & s & ". Reinstall Palms LAN Center to put them back."
    ElseIf Len(gPakNote) > 0 And gPakCount > 0 Then
        ic = IC_WARN
        s = gPakNote
    Else
        ic = IC_HDD
        s = "Games are installed to " & GamesRoot()
        free = FreeBytes(GamesRoot())
        If free >= 0 Then s = s & " (" & SizeText(free) & " free)"
        s = s & ". This can be changed in Tools > Options."
    End If
    DrawIcon16 picLibPage, ic, 18, h - 10 - 34 + 9
    TextAt picLibPage, 40, h - 10 - 34 + 10, FitText(picLibPage, s, w - 60), vbBlack
    DrawLibListNow
    DrawLibCardNow
    UpdateLibButtons
End Sub

' ---- the list

Private Function LibShown(ByVal id As String) As Boolean
    Dim i As Long
    If Len(id) = 0 Then Exit Function
    For i = 0 To gShowCount - 1
        If gShowId(i) = id Then
            LibShown = True
            Exit Function
        End If
    Next
End Function

' Installed games first, then the rest, each in name order, with a heading over each group
Private Sub BuildLibRows()
    Dim i As Long, n As Long, pass As Integer, idx() As Long, keys() As String
    ReDim mLibRowKind(0 To gShowCount + 2)
    ReDim mLibRowId(0 To gShowCount + 2)
    mLibRowCount = 0
    mLibHereN = 0
    mLibNewN = 0
    If gShowCount = 0 Then Exit Sub
    ReDim idx(0 To gShowCount - 1)
    ReDim keys(0 To gShowCount - 1)
    For pass = 0 To 1
        n = 0
        For i = 0 To gShowCount - 1
            If LibInstalled(gShowId(i)) = (pass = 0) Then
                idx(n) = i
                keys(n) = LCase$(LibName(gShowId(i)))
                n = n + 1
            End If
        Next
        If n > 0 Then
            If pass = 0 Then
                mLibHereN = n
                AddLibRow 0, "here"
            Else
                mLibNewN = n
                AddLibRow 0, "new"
            End If
            SortLib idx, keys, n
            For i = 0 To n - 1
                AddLibRow 1, gShowId(idx(i))
            Next
        End If
    Next
End Sub

Private Sub AddLibRow(ByVal kind As Integer, ByVal id As String)
    mLibRowKind(mLibRowCount) = kind
    mLibRowId(mLibRowCount) = id
    mLibRowCount = mLibRowCount + 1
End Sub

Private Sub SortLib(idx() As Long, keys() As String, ByVal n As Long)
    Dim i As Long, j As Long, ti As Long, tk As String
    For i = 1 To n - 1
        j = i
        Do While j > 0
            If keys(j - 1) <= keys(j) Then Exit Do
            ti = idx(j - 1)
            idx(j - 1) = idx(j)
            idx(j) = ti
            tk = keys(j - 1)
            keys(j - 1) = keys(j)
            keys(j) = tk
            j = j - 1
        Loop
    Next
End Sub

Private Sub CalcLibCols(ByVal w As Long)
    Dim nameW As Long
    nameW = w - LIB_PLAYERS_W - LIB_SIZE_W - LIB_STATUS_W
    If nameW < 100 Then nameW = 100
    mLibColX(0) = 0
    mLibColX(1) = nameW
    mLibColX(2) = mLibColX(1) + LIB_PLAYERS_W
    mLibColX(3) = mLibColX(2) + LIB_SIZE_W
    mLibColX(4) = w
    If mLibColX(4) < mLibColX(3) + 40 Then mLibColX(4) = mLibColX(3) + 40
End Sub

Private Function LibColName(ByVal c As Integer) As String
    Select Case c
    Case 0: LibColName = "Game"
    Case 1: LibColName = "Players"
    Case 2: LibColName = "Size"
    Case Else: LibColName = "Status"
    End Select
End Function

Private Function LibVisible() As Long
    LibVisible = (picLibList.ScaleHeight - HDR_H - 1) \ LIB_ROW
    If LibVisible < 1 Then LibVisible = 1
End Function

Private Sub UpdateLibScroll()
    Dim vis As Long, mx As Long
    vis = LibVisible()
    mx = mLibRowCount - vis
    If mx < 0 Then mx = 0
    mIgnoreScroll = True
    If vsbLib.Value > mx Then vsbLib.Value = mx
    vsbLib.Max = mx
    If vis > 1 Then
        vsbLib.LargeChange = vis - 1
    Else
        vsbLib.LargeChange = 1
    End If
    vsbLib.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

Private Sub DrawLibList()
    Want DR_LIBLIST
End Sub

Private Sub DrawLibListNow()
    Dim w As Long, c As Integer, r As Long, y As Long, s As String, tw As Long, id As String
    Dim st As String, col As Long, ic As Integer, pct As Long, slot As Long, nameCol As Long, subCol As Long, sw As Long
    picLibList.Cls
    w = picLibList.ScaleWidth - 16
    CalcLibCols w
    For c = 0 To 3
        DrawLibHead c
    Next
    If mLibRowCount = 0 Then
        TextAt picLibList, 8, HDR_H + 8, FitText(picLibList, LibEmptyText(), w - 16), C_GREY
        Exit Sub
    End If
    For r = vsbLib.Value To mLibRowCount - 1
        y = HDR_H + 1 + (r - vsbLib.Value) * LIB_ROW
        If y > picLibList.ScaleHeight Then Exit For
        If mLibRowKind(r) = 0 Then
            ' a group heading, with a line after it
            If mLibRowId(r) = "here" Then
                s = "Installed (" & mLibHereN & ")"
            Else
                s = "Not Installed (" & mLibNewN & ")"
            End If
            TextAt picLibList, 6, y + 4, s, C_TITLE, True
            picLibList.FontBold = True
            tw = picLibList.TextWidth(s)
            picLibList.FontBold = False
            Box picLibList, 6 + tw + 6, y + 10, w - (6 + tw + 6) - 6, 1, vb3DShadow
        Else
            id = mLibRowId(r)
            LibStatus id, st, col, ic, pct
            nameCol = vbBlack
            subCol = C_GREY
            If id = mLibSel Then
                Box picLibList, 1, y, w - 2, LIB_ROW, C_SEL
                nameCol = vbWhite
                subCol = vbWhite
                col = vbWhite
            End If
            slot = GameIconSlot(LibIconPath(id))
            If slot >= 0 Then
                DrawGameIcon picLibList, slot, False, 8, y + 2
            Else
                DrawIcon16 picLibList, IC_DISC, 8, y + 2
            End If
            ' the name, then the version in grey if there's room
            s = FitText(picLibList, LibName(id), mLibColX(1) - 34)
            TextAt picLibList, 28, y + 4, s, nameCol
            tw = picLibList.TextWidth(s)
            s = LibVer(id)
            If Len(s) > 0 Then
                sw = mLibColX(1) - 34 - tw - 6
                If sw > 24 Then TextAt picLibList, 28 + tw + 6, y + 4, FitText(picLibList, s, sw), subCol
            End If
            TextAt picLibList, mLibColX(1) + 6, y + 4, FitText(picLibList, LibPlayers(id), mLibColX(2) - mLibColX(1) - 8), nameCol
            TextAt picLibList, mLibColX(2) + 6, y + 4, FitText(picLibList, SizeText(LibSize(id)), mLibColX(3) - mLibColX(2) - 8), nameCol
            ' what it's up to: a progress bar while it installs, otherwise an icon and words
            If pct >= 0 Then
                ProgressBar picLibList, mLibColX(3) + 6, y + 4, 44, 12, pct
                TextAt picLibList, mLibColX(3) + 56, y + 4, FitText(picLibList, st, mLibColX(4) - mLibColX(3) - 60), col
            Else
                DrawIcon16 picLibList, ic, mLibColX(3) + 4, y + 2
                TextAt picLibList, mLibColX(3) + 24, y + 4, FitText(picLibList, st, mLibColX(4) - mLibColX(3) - 28), col
            End If
        End If
    Next
End Sub

' What the list says when there's nothing in it
Private Function LibEmptyText() As String
    If mSignedIn And Not ServerHasPaks() Then
        LibEmptyText = "The Game Library requires Palms Server version 0.4.3 or later."
    ElseIf Not gPakHeard Then
        LibEmptyText = "Retrieving the cartridge list..."
    ElseIf Len(gPakNote) > 0 Then
        LibEmptyText = gPakNote
    Else
        LibEmptyText = "No cartridges available. To add a game, click Make Cartridge on the toolbar."
    End If
End Function

' A column heading, a raised button like Explorer's
Private Sub DrawLibHead(ByVal c As Integer)
    Dim x As Long, w As Long
    x = mLibColX(c)
    w = mLibColX(c + 1) - x
    Box picLibList, x, 0, w, HDR_H, vbButtonFace
    Edge picLibList, x, 0, w, HDR_H, vb3DHighlight, vb3DDKShadow
    Box picLibList, x + 1, HDR_H - 2, w - 2, 1, vb3DShadow
    Box picLibList, x + w - 2, 1, 1, HDR_H - 2, vb3DShadow
    TextAt picLibList, x + 6, 3, FitText(picLibList, LibColName(c), w - 12), vbButtonText
End Sub

Private Function LibRowAt(ByVal py As Single) As Long
    Dim r As Long
    LibRowAt = -1
    If py < HDR_H + 1 Then Exit Function
    r = vsbLib.Value + Int((py - HDR_H - 1) / LIB_ROW)
    If r < mLibRowCount Then LibRowAt = r
End Function

Private Sub SelectGame(ByVal id As String)
    If id = mLibSel Then Exit Sub
    mLibSel = id
    SettingSet "lib_sel", id
    DrawLibPage
    RefreshLibTools
End Sub

Private Sub picLibList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Long
    On Error Resume Next
    picLibList.SetFocus
    On Error GoTo 0
    mLibDownRow = -1
    r = LibRowAt(Y)
    If r < 0 Then Exit Sub
    If mLibRowKind(r) = 0 Then Exit Sub
    mLibDownRow = r
    SelectGame mLibRowId(r)
End Sub

' Double-click: install it, or play it
Private Sub picLibList_DblClick()
    If mLibDownRow < 0 Then Exit Sub
    LibGo
End Sub

' Up and down move through the list, Enter installs or plays
Private Sub picLibList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long, cur As Long, stp As Long
    cur = -1
    For r = 0 To mLibRowCount - 1
        If mLibRowKind(r) <> 0 And mLibRowId(r) = mLibSel Then cur = r
    Next
    Select Case KeyCode
    Case vbKeyUp
        stp = -1
    Case vbKeyDown
        stp = 1
    Case vbKeyReturn
        LibGo
        Exit Sub
    Case vbKeyDelete
        LibDel
        Exit Sub
    Case Else
        Exit Sub
    End Select
    r = cur + stp
    Do While r >= 0 And r < mLibRowCount
        If mLibRowKind(r) <> 0 Then
            If r < vsbLib.Value Then vsbLib.Value = r
            If r >= vsbLib.Value + LibVisible() Then vsbLib.Value = r - LibVisible() + 1
            SelectGame mLibRowId(r)
            Exit Sub
        End If
        r = r + stp
    Loop
End Sub

Private Sub vsbLib_Change()
    If Not mIgnoreScroll Then DrawLibList
End Sub

Private Sub vsbLib_Scroll()
    DrawLibList
End Sub

' ---- the card about the game picked

Private Sub DrawLibCard()
    Want DR_LIBCARD
End Sub

Private Sub DrawLibCardNow()
    Dim w As Long, y As Long, id As String, s As String, st As String, col As Long, ic As Integer, pct As Long
    Dim slot As Long, c As Long, h As Long, state As Integer
    picLibCard.Cls
    w = picLibCard.ScaleWidth
    id = mLibSel
    If Not LibShown(id) Then
        DrawIcon32 picLibCard, IC_DISC, (w - 32) \ 2, 30
        s = "Select a game in the list"
        TextAt picLibCard, (w - picLibCard.TextWidth(s)) \ 2, 72, s, C_SUB
        s = "to view details."
        TextAt picLibCard, (w - picLibCard.TextWidth(s)) \ 2, 88, s, C_SUB
        Exit Sub
    End If
    c = CatIndex(PakOf(id))
    h = HereIndex(PakOf(id))
    state = LibState(id)
    slot = GameIconSlot(LibIconPath(id))
    If slot >= 0 Then
        DrawGameIcon picLibCard, slot, True, 4, 4
    Else
        DrawIcon32 picLibCard, IC_DISC, 4, 4
    End If
    picLibCard.FontSize = 12
    TextAt picLibCard, 44, 2, FitText(picLibCard, LibName(id), w - 48), C_TITLE, True
    picLibCard.FontSize = 8
    s = LibVer(id)
    If Len(LibPlayers(id)) > 0 Then
        If Len(s) > 0 Then s = s & ", "
        s = s & LibPlayers(id) & " players"
    End If
    TextAt picLibCard, 45, 24, FitText(picLibCard, s, w - 48), C_SUB
    y = 44
    Box picLibCard, 2, y, w - 4, 1, vb3DShadow
    Box picLibCard, 2, y + 1, w - 4, 1, vb3DHighlight
    y = y + 7

    ' what it's up to
    LibStatus id, st, col, ic, pct
    DrawIcon16 picLibCard, ic, 4, y
    Select Case state
    Case LS_BUSY
        y = y + 2
        WrapLines picLibCard, gJobText, 24, y, w - 28, C_BLUE, True, 2
        y = y + 3
        If pct >= 0 Then
            ProgressBar picLibCard, 6, y, w - 12, 14, pct
            y = y + 20
        End If
    Case LS_WAIT
        y = y + 2
        WrapLines picLibCard, "Queued. Waiting for the current installation to finish.", 24, y, w - 28, C_GREY, True, 2
        y = y + 3
    Case LS_UPDATE
        TextAt picLibCard, 24, y + 2, "Update available", C_ORANGE, True
        y = y + 19
        s = VerText(gPak(c).Ver, gPak(c).Rev)
        If Len(s) > 0 Then s = s & ", "
        WrapLines picLibCard, "A newer cartridge (" & s & "made " & CartDay(gPak(c).Made) & ") is on the Cartridge Server. Click Update to install it.", 6, y, w - 12, C_SUB, False, 3
    Case LS_OLD
        TextAt picLibCard, 24, y + 2, "Installed", C_GREEN, True
        y = y + 19
        WrapLines picLibCard, "No longer available on the Cartridge Server. The installed copy can still be played.", 6, y, w - 12, C_SUB, False, 2
    Case LS_READY
        If LibFailed(id) Then
            TextAt picLibCard, 24, y + 2, "Installed with errors", C_ORANGE, True
        Else
            TextAt picLibCard, 24, y + 2, "Installed", C_GREEN, True
        End If
        y = y + 19
    Case Else
        TextAt picLibCard, 24, y + 2, st, col, True
        y = y + 19
    End Select
    If LibFailed(id) And Len(gFailText) > 0 Then
        WrapLines picLibCard, gFailText, 6, y, w - 12, C_RED, False, 4
    End If
    y = y + 4

    ' the details
    LibCardRow y, "Players", LibPlayers(id), vbBlack
    LibCardRow y, "Size", SizeText(LibSize(id)), vbBlack
    If c >= 0 And (state = LS_NONE Or state = LS_UPDATE Or state = LS_WAIT Or state = LS_BUSY) Then
        If gPak(c).ArcBytes > 0 Then LibCardRow y, "Download", SizeText(gPak(c).ArcBytes), vbBlack
    End If
    If LibPackCount(id) <= 1 Then LibCardRow y, "Cartridge", CartText(id), vbBlack
    LibCardRow y, "Folder", LibPlace(id), vbBlack
    If h >= 0 Then LibCardRow y, "Installed", NiceWhen(gHere(h).Installed), vbBlack
    If h >= 0 Then
        If LibNotesUnread(id) Then
            LibCardRow y, "Setup notes", "Yes (unread)", C_ORANGE
        ElseIf LibHasNotes(id) Then
            LibCardRow y, "Setup notes", "Yes", vbBlack
        End If
    ElseIf c >= 0 Then
        If Len(gPak(c).Notes) > 0 Then LibCardRow y, "Setup notes", "Yes (shown after installation)", vbBlack
    End If

    ' a game in a pack comes with the others
    If LibPackCount(id) > 1 Then
        LibCardRow y, "Multicart", LibPakName(id) & ", " & CartText(id), vbBlack
        y = y + 4
        WrapLines picLibCard, "Includes " & LibPackOthers(id) & ". These games are installed, updated and uninstalled together.", 6, y, w - 12, C_SUB, False, 4
    End If
End Sub

Private Sub LibCardRow(y As Long, ByVal lbl As String, ByVal txt As String, ByVal col As Long)
    If Len(txt) = 0 Then Exit Sub
    TextAt picLibCard, 6, y, lbl, C_SUB
    TextAt picLibCard, 78, y, FitText(picLibCard, txt, picLibCard.ScaleWidth - 82), col
    y = y + 14
End Sub

' Writes words over as many lines as they need, up to maxLines (the last one ends in ...).
' y moves down past them.
Private Sub WrapLines(p As PictureBox, ByVal s As String, ByVal x As Long, y As Long, ByVal w As Long, ByVal col As Long, ByVal bold As Boolean, ByVal maxLines As Long)
    Dim words() As String, i As Long, k As Long, cur As String, try As String, n As Long, rest As String
    words = Split(Trim$(s), " ")
    p.FontBold = bold
    For i = 0 To UBound(words)
        If Len(cur) = 0 Then
            try = words(i)
        Else
            try = cur & " " & words(i)
        End If
        If p.TextWidth(try) > w And Len(cur) > 0 Then
            n = n + 1
            If n >= maxLines Then
                ' out of lines: the rest goes on this one, cut short
                rest = cur
                For k = i To UBound(words)
                    rest = rest & " " & words(k)
                Next
                TextAt p, x, y, FitText(p, rest, w), col, bold
                y = y + 14
                p.FontBold = False
                Exit Sub
            End If
            TextAt p, x, y, cur, col, bold
            p.FontBold = bold
            y = y + 14
            cur = words(i)
        Else
            cur = try
        End If
    Next
    If Len(cur) > 0 Then
        TextAt p, x, y, FitText(p, cur, w), col, bold
        y = y + 14
    End If
    p.FontBold = False
End Sub

' "2026-10-22 18:40" becomes "22 Oct, 6:40 PM"
Private Function NiceWhen(ByVal s As String) As String
    Dim d As Date
    On Error GoTo Plain
    d = CDate(s)
    NiceWhen = Format$(d, "d mmm") & ", " & AmPm(Format$(d, "hh:nn"))
    Exit Function
Plain:
    NiceWhen = s
End Function

' The cartridge a game comes on: its version and the day it was made, like v1.1.1.0, made 23 Oct 2026
Private Function CartText(ByVal id As String) As String
    Dim c As Long, h As Long, made As String
    h = HereIndex(PakOf(id))
    c = CatIndex(PakOf(id))
    If h >= 0 Then
        made = gHere(h).Made
    ElseIf c >= 0 Then
        made = gPak(c).Made
    End If
    CartText = LibVer(id)
    If Len(made) > 0 Then
        If Len(CartText) > 0 Then CartText = CartText & ", "
        CartText = CartText & "made " & CartDay(made)
    End If
End Function

Private Function CartDay(ByVal s As String) As String
    Dim d As Date
    On Error GoTo Plain
    d = CDate(s)
    CartDay = Format$(d, "d mmm yyyy")
    Exit Function
Plain:
    CartDay = s
End Function

' ---- the buttons on the card, and the toolbar

' Only touches a button when something about it changes, so it doesn't flicker
' every time the progress bar moves
Private Sub UpdateLibButtons()
    Dim st As Integer, shown As Boolean, goText As String, altText As String, goOn As Boolean, delOn As Boolean
    shown = LibShown(mLibSel)
    st = LibState(mLibSel)
    goOn = True
    If shown Then
        Select Case st
        Case LS_NONE
            goText = "&Install"
            goOn = LibCanInstall(mLibSel)
        Case LS_READY, LS_OLD, LS_UPDATE
            delOn = True
            If Not LibInstalled(mLibSel) Then
                ' a game the installed copy of its pack hasn't got: updating the pack gets it
                goText = "&Update"
            Else
                goText = "&Play"
                If st = LS_UPDATE Then
                    altText = "&Update"
                ElseIf LibHasSettings(mLibSel) Then
                    altText = "&Configure"
                End If
            End If
        Case LS_BUSY, LS_WAIT
            goText = "&Cancel"
        End Select
    End If
    If cmdLibGo.Caption <> goText And Len(goText) > 0 Then cmdLibGo.Caption = goText
    If cmdLibGo.Enabled <> goOn Then cmdLibGo.Enabled = goOn
    If cmdLibGo.Visible <> shown Then cmdLibGo.Visible = shown
    If cmdLibAlt.Caption <> altText And Len(altText) > 0 Then cmdLibAlt.Caption = altText
    If cmdLibAlt.Visible <> (Len(altText) > 0) Then cmdLibAlt.Visible = (Len(altText) > 0)
    If cmdLibDel.Visible <> delOn Then cmdLibDel.Visible = delOn
End Sub

Private Sub BuildLibTools()
    Dim st As Integer
    If LibShown(mLibSel) Then
        st = LibState(mLibSel)
        Select Case st
        Case LS_NONE
            If LibCanInstall(mLibSel) Then AddTool "lgo", IC_DOWNLOAD, "Install", False
        Case LS_READY, LS_OLD
            AddTool "lgo", IC_JOYSTICK, "Play", False
        Case LS_UPDATE
            If LibInstalled(mLibSel) Then AddTool "lgo", IC_JOYSTICK, "Play", False
            AddTool "lupd", IC_REFRESH, "Update", False
        Case LS_BUSY, LS_WAIT
            AddTool "lgo", IC_BAD, "Cancel", False
        End Select
        If st = LS_READY Or st = LS_OLD Or st = LS_UPDATE Then
            If LibHasSettings(mLibSel) Then AddTool "lcfg", IC_PENCIL, "Configure", False
            If LibHasNotes(mLibSel) Then AddTool "lnotes", IC_INFO, "Setup Notes", False
            AddTool "lfolder", IC_FOLDER, "Open Folder", False
            AddTool "-", -1, "", False
            AddTool "ldel", IC_ERASER, "Uninstall", False
        End If
        AddTool "-", -1, "", False
    End If
    AddTool "lroot", IC_HDD, "Games Folder...", False
    AddTool "-", -1, "", False
    AddTool "cart", IC_BOX, "Make Cartridge...", False
End Sub

Private Sub RefreshLibTools()
    If mPage = PG_LIBRARY Then
        BuildTools
        DrawTools
    End If
End Sub

' The big button: install, play, or stop installing
Private Sub LibGo()
    Select Case LibState(mLibSel)
    Case LS_NONE
        If LibAskInstall(mLibSel) Then LibRefresh
    Case LS_READY, LS_OLD, LS_UPDATE
        If LibInstalled(mLibSel) Then
            LibPlay mLibSel, Me.hWnd, False
            DrawLibCard
        ElseIf LibAskInstall(mLibSel) Then
            ' the installed copy of its pack hasn't got it: the button says Update
            LibRefresh
        End If
    Case LS_BUSY, LS_WAIT
        If LibAskCancel(mLibSel) Then LibRefresh
    End Select
End Sub

' The other button: update, or the game's settings program
Private Sub LibAlt()
    If LibState(mLibSel) = LS_UPDATE Then
        If LibAskInstall(mLibSel) Then LibRefresh
    ElseIf LibHasSettings(mLibSel) Then
        LibPlay mLibSel, Me.hWnd, True
    End If
End Sub

Private Sub LibDel()
    If LibAskUninstall(mLibSel) Then LibRefresh
End Sub

Private Sub cmdLibGo_Click()
    LibGo
End Sub

Private Sub cmdLibAlt_Click()
    LibAlt
End Sub

Private Sub cmdLibDel_Click()
    LibDel
End Sub

' ---- keeping an eye on installs

' modLibrary has work on: look in on it twice a second
Public Sub LibWake()
    If Not tmrLib.Enabled Then tmrLib.Enabled = True
End Sub

Private Sub tmrLib_Timer()
    Dim changed As Boolean
    If mInLibTick Then Exit Sub
    mInLibTick = True
    PerfStart PF_LIBTICK
    changed = LibTick()
    PerfStop PF_LIBTICK
    If gIconsNew Then
        gIconsNew = False
        GameIconsRetry
        DrawLibPage
    End If
    If gLibDirty Then
        LibRefresh
    ElseIf changed Then
        Want DR_LIBLIST Or DR_LIBCARD
        LibNavBadge
    End If
    If Not LibWorking() Then tmrLib.Enabled = False
    mInLibTick = False
End Sub

' The little blue bubble on Game Library in the sidebar while a game installs
Private Function LibBadgeText() As String
    If Len(gJobId) = 0 Then Exit Function
    If gJobPct >= 0 Then
        LibBadgeText = gJobPct & "%"
    Else
        LibBadgeText = "..."
    End If
End Function

Private Sub LibNavBadge()
    Dim s As String
    s = LibBadgeText()
    If s <> mLibBadge Then
        mLibBadge = s
        DrawSide
    End If
End Sub

' modLibrary: a game finished installing or uninstalling, or something went wrong
' (id is the pak's id: for a pack, all of its games were installed or uninstalled)
' Add/Remove Programs started the program with /uninstall <cartridge> (see Sub Main): ask, and uninstall it
Public Sub UninstallFromCommand(ByVal id As String)
    Dim key As String
    id = LCase$(Trim$(id))
    If HereIndex(id) < 0 Then
        ' an entry left behind by a game that's gone already: tidy it away
        DropUninstallEntry id
        MsgBox "This game is not installed on this PC by Palms LAN Center. It has been removed from Add/Remove Programs.", vbInformation, "Palms LAN Center"
        Exit Sub
    End If
    key = LibFirstGame(id)
    If Len(key) = 0 Then key = id & "/" & id
    If LibAskUninstall(key) Then
        mLibSel = key
        LibRefresh
    End If
End Sub

Public Sub LibFinished(ByVal id As String, ByVal ok As Boolean, ByVal msg As String, ByVal hasNotes As Boolean)
    Dim nm As String
    nm = LibPakName(id)
    LibRefresh
    If Not ok Then
        Note nm & ": " & msg
        Ding
        ShowToast "lib", IC_WARN, nm & " - Error", msg, 12
    ElseIf LibPakInstalled(id) Then
        Note msg
        Ding
        If id = gFailId Then
            ' installed, but something needs sorting out (like its registry settings)
            ShowToast "lib", IC_WARN, nm & " installed with errors", gFailText, 15
        ElseIf hasNotes Then
            mLibToastId = LibFirstGame(id)
            ShowToast "libnotes", IC_OK, nm & " installed", "This game has setup notes. Click here to view them.", 15
        Else
            ShowToast "lib", IC_OK, nm & " installed", "The game can be started from the Game Library.", 8
        End If
    Else
        Note msg
    End If
End Sub

' ================================================================ What's On

Private Sub LayoutWhats()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long
    w = picWhatsPage.ScaleWidth
    h = picWhatsPage.ScaleHeight
    lw = ((w - 30) * 11) \ 20
    rx = 10 + lw + 10
    rw = w - rx - 10
    gH = h - 16 - 10 - 34 - 10
    If gH < 80 Then gH = 80
    picSched.Move 16, 28, lw - 12, gH - 18
    vsbSched.Move picSched.ScaleWidth - 16, 0, 16, picSched.ScaleHeight
    picVote.Move rx + 6, 28, rw - 12, gH - 18
    DrawWhats
End Sub

' The group boxes and tip around the two lists, then the lists themselves
Private Sub DrawWhats()
    Want DR_WHATS
End Sub

Private Sub DrawWhatsNow()
    Dim w As Long, h As Long, lw As Long, rx As Long, rw As Long, gH As Long, s As String
    w = picWhatsPage.ScaleWidth
    h = picWhatsPage.ScaleHeight
    lw = ((w - 30) * 11) \ 20
    rx = 10 + lw + 10
    rw = w - rx - 10
    gH = h - 16 - 10 - 34 - 10
    If gH < 80 Then gH = 80
    picWhatsPage.Cls
    GroupBox picWhatsPage, 10, 16, lw, gH, IC_CALENDAR, "Today's Schedule"
    GroupBox picWhatsPage, rx, 16, rw, gH, IC_STAR, "Next Game Vote"
    TipBox picWhatsPage, 10, h - 10 - 34, w - 20, 34
    DrawIcon16 picWhatsPage, IC_INFO, 18, h - 10 - 34 + 9
    If mIsAdmin Then
        s = "Use the toolbar to add events or start a vote. Right-click an event for more options."
    Else
        s = "Select an event for details and reminders. To vote, click a game on the right."
    End If
    TextAt picWhatsPage, 40, h - 10 - 34 + 10, FitText(picWhatsPage, s, w - 60), vbBlack
    UpdateSchedScroll
    DrawSchedNow
    DrawVoteNow
End Sub

' The event that's on now: the last one that has started. -1 if none has.
Private Function CurrentEvent() As Long
    Dim i As Long, nowKey As Long
    CurrentEvent = -1
    nowKey = PartyMinutes(Format$(Now, "hh:nn"))
    For i = 0 To mSchedCount - 1
        If PartyMinutes(mSchedTime(i)) <= nowKey Then CurrentEvent = i
    Next
End Function

Private Sub UpdateSchedScroll()
    Dim vis As Long, mx As Long
    vis = (picSched.ScaleHeight - 4) \ SCHED_ROW
    If vis < 1 Then vis = 1
    mx = mSchedCount - vis
    If mx < 0 Then mx = 0
    mIgnoreScroll = True
    If vsbSched.Value > mx Then vsbSched.Value = mx
    vsbSched.Max = mx
    If vis > 1 Then
        vsbSched.LargeChange = vis - 1
    Else
        vsbSched.LargeChange = 1
    End If
    vsbSched.Enabled = (mx > 0)
    mIgnoreScroll = False
End Sub

Private Sub DrawSched()
    Want DR_SCHED
End Sub

Private Sub DrawSchedNow()
    Dim w As Long, i As Long, y As Long, cur As Long, s As String, tw As Long, rx As Long
    Dim tcol As Long, lcol As Long
    picSched.Cls
    w = picSched.ScaleWidth - 16
    If mSchedCount = 0 Then
        s = "No events scheduled."
        TextAt picSched, (w - picSched.TextWidth(s)) \ 2, 30, s, C_GREY
        If mIsAdmin Then
            s = "Click Add Event on the toolbar."
            TextAt picSched, (w - picSched.TextWidth(s)) \ 2, 48, s, C_GREY
        End If
        Exit Sub
    End If
    cur = CurrentEvent()
    For i = vsbSched.Value To mSchedCount - 1
        y = 2 + (i - vsbSched.Value) * SCHED_ROW
        If y > picSched.ScaleHeight Then Exit For
        tcol = vbBlack
        If i = mSchedSel Then
            Box picSched, 2, y, w - 4, SCHED_ROW, C_SEL
            tcol = vbWhite
        ElseIf i = cur Then
            Box picSched, 2, y, w - 4, SCHED_ROW, C_TIP
        ElseIf i < cur Then
            tcol = C_GREY
        End If
        ' time, right-lined-up, then the icon and what's on
        s = AmPm(mSchedTime(i))
        picSched.FontBold = True
        tw = picSched.TextWidth(s)
        picSched.FontBold = False
        TextAt picSched, 64 - tw, y + 5, s, tcol, True
        DrawIcon16 picSched, mSchedIcon(i), 72, y + 4
        ' "On now" or "Up next" on the right
        s = ""
        If i = cur Then
            s = "Now"
            lcol = C_GREEN
        ElseIf i = cur + 1 Then
            s = "Next"
            lcol = C_BLUE
        End If
        If i = mSchedSel Then lcol = vbWhite
        rx = w - 8
        If Len(s) > 0 Then
            picSched.FontBold = True
            tw = picSched.TextWidth(s)
            picSched.FontBold = False
            rx = rx - tw
            TextAt picSched, rx, y + 5, s, lcol, True
            rx = rx - 8
        End If
        ' then, right to left: a bell for Remind Me, an i if there are details, and how many are in
        If IsRemind(i) Then
            rx = rx - 16
            DrawIcon16 picSched, IC_BELL, rx, y + 4
            rx = rx - 4
        End If
        If Len(mSchedNote(i)) > 0 Then
            rx = rx - 16
            DrawIcon16 picSched, IC_INFO, rx, y + 4
            rx = rx - 4
        End If
        If mSchedInN(i) > 0 Then
            s = mSchedInN(i) & " in"
            If IsInEvent(i) Then s = s & " (you)"
            If i = mSchedSel Then
                lcol = vbWhite
            ElseIf IsInEvent(i) Then
                lcol = C_GREEN
            Else
                lcol = C_SUB
            End If
            rx = rx - picSched.TextWidth(s)
            TextAt picSched, rx, y + 5, s, lcol
            rx = rx - 8
        End If
        TextAt picSched, 94, y + 5, FitText(picSched, mSchedName(i), rx - 94 - 4), tcol
    Next
End Sub

Private Sub DrawVote()
    Want DR_VOTE
End Sub

Private Sub DrawVoteNow()
    Dim w As Long, i As Long, y As Long, s As String, tw As Long, bw As Long, fillW As Long, x As Long, ic As Integer
    picVote.Cls
    w = picVote.ScaleWidth
    If mVoteState <> "open" And mVoteState <> "closed" Then
        DrawIcon32 picVote, IC_STAR, (w - 32) \ 2, 30
        s = "No vote in progress."
        TextAt picVote, (w - picVote.TextWidth(s)) \ 2, 72, s, C_SUB
        s = "Votes started by the host are shown here."
        TextAt picVote, (w - picVote.TextWidth(s)) \ 2, 90, FitText(picVote, s, w - 8), C_GREY
        Exit Sub
    End If
    picVote.FontBold = True
    s = FitText(picVote, mVoteQ, w - 20)
    picVote.FontBold = False
    TextAt picVote, 10, 8, s, vbBlack, True
    If mVoteState = "open" Then
        s = "Voting closes in " & MmSs(VoteSecsLeft()) & "."
        If mVoteMine = 0 Then
            s = s & " Click a game to vote."
        Else
            s = s & " Your vote can be changed."
        End If
        TextAt picVote, 10, 26, FitText(picVote, s, w - 20), C_BLUE
    Else
        If mVoteWinner > 0 Then
            s = "Voting closed. Winner: " & mVoteName(mVoteWinner - 1) & "."
        ElseIf mVoteTotal = 0 Then
            s = "Voting closed. No votes were cast."
        Else
            s = "Voting closed. The result is a tie."
        End If
        TextAt picVote, 10, 26, FitText(picVote, s, w - 20), C_GREEN
    End If
    For i = 0 To mVoteCount - 1
        y = VOTE_TOP + i * VOTE_ROW
        If mVoteState = "closed" And mVoteWinner = i + 1 Then Box picVote, 4, y - 2, w - 8, VOTE_ROW - 2, C_TIP
        If mVoteMine = i + 1 Then
            ic = IC_RADIO1
        Else
            ic = IC_RADIO0
        End If
        DrawIcon16 picVote, ic, 10, y + 1
        If mVoteVotes(i) = 1 Then
            s = "1 vote"
        Else
            s = mVoteVotes(i) & " votes"
        End If
        tw = picVote.TextWidth(s)
        TextAt picVote, w - 10 - tw, y + 3, s, C_SUB
        TextAt picVote, 32, y + 3, FitText(picVote, mVoteName(i), w - 32 - tw - 20), vbBlack, (mVoteWinner = i + 1)
        ' a bar showing the share of votes, in blocks like a Windows progress bar
        bw = w - 32 - 10
        Edge picVote, 32, y + 19, bw, 11, vb3DShadow, vb3DHighlight
        fillW = 0
        If mVoteTotal > 0 Then fillW = ((bw - 4) * mVoteVotes(i)) \ mVoteTotal
        x = 0
        Do While x + 7 <= fillW
            Box picVote, 34 + x, y + 21, 7, 7, C_SEL
            x = x + 9
        Loop
    Next
End Sub

' The pointing hand over the choices while voting is open
Private Sub picVote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Long
    If mVoteState <> "open" Or Y < VOTE_TOP - 2 Then Exit Sub
    n = Int((Y - VOTE_TOP + 2) / VOTE_ROW) + 1
    If n >= 1 And n <= mVoteCount Then HandCursor
End Sub

' Click a choice to vote for it
Private Sub picVote_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Long
    If Button <> vbLeftButton Then Exit Sub
    If mVoteState <> "open" Then Exit Sub
    If Y < VOTE_TOP - 2 Then Exit Sub
    n = Int((Y - VOTE_TOP + 2) / VOTE_ROW) + 1
    If n < 1 Or n > mVoteCount Then Exit Sub
    If n = mVoteMine Then Exit Sub
    mVoteMine = n
    DrawVote
    SendLine "VOTE|" & n
End Sub

' Click an event to select it: the toolbar then has its buttons (Count Me In, Remind Me...)
Private Sub picSched_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Long
    n = vsbSched.Value + Int((Y - 2) / SCHED_ROW)
    If n < 0 Or n >= mSchedCount Then n = -1
    If n <> mSchedSel Then
        mSchedSel = n
        DrawSched
        RefreshWhatsTools
    End If
End Sub

' Right-click an event for its menu
Private Sub picSched_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And mSchedSel >= 0 Then ShowEventMenu
End Sub

' Double-click an event for its details
Private Sub picSched_DblClick()
    If mSchedSel >= 0 Then ShowEventInfo mSchedSel
End Sub

' The What's On toolbar changes with the selection and the vote, so rebuild it
Private Sub RefreshWhatsTools()
    If mPage = PG_WHATSON Then
        BuildTools
        DrawTools
    End If
End Sub

Private Sub RemoveSelectedEvent()
    If mSchedSel < 0 Or mSchedSel >= mSchedCount Then
        Note "Select an event in Today's Schedule first."
        Exit Sub
    End If
    If MsgBox("Remove " & AmPm(mSchedTime(mSchedSel)) & " " & mSchedName(mSchedSel) & " from the schedule?", vbQuestion + vbOKCancel, "Remove Event") = vbOK Then
        SendLine "SCHEDDEL|" & mSchedId(mSchedSel)
    End If
End Sub

Private Sub vsbSched_Change()
    If Not mIgnoreScroll Then DrawSched
End Sub

Private Sub vsbSched_Scroll()
    DrawSched
End Sub

' Turns the SCHED lines we collected into the list (the server sends them in time order)
Private Sub FinishSched()
    Dim n As Long, i As Long, k As Long, f() As String, v As Variant, selId As String, keep As String
    If mSchedSel >= 0 And mSchedSel < mSchedCount Then selId = mSchedId(mSchedSel)
    n = 0
    If Not mSchedBuild Is Nothing Then n = mSchedBuild.Count
    mSchedSel = -1
    If n > 0 Then
        ReDim mSchedId(0 To n - 1)
        ReDim mSchedTime(0 To n - 1)
        ReDim mSchedIcon(0 To n - 1)
        ReDim mSchedName(0 To n - 1)
        ReDim mSchedIn(0 To n - 1)
        ReDim mSchedInN(0 To n - 1)
        ReDim mSchedNote(0 To n - 1)
        ReDim mSchedGame(0 To n - 1)
        ReDim mSchedServer(0 To n - 1)
        ReDim mSchedConnect(0 To n - 1)
        i = 0
        For Each v In mSchedBuild
            f = Split(CStr(v), "|")
            mSchedId(i) = Fld(f, 1)
            mSchedTime(i) = Fld(f, 2)
            mSchedIcon(i) = EventIcon(Fld(f, 3))
            mSchedName(i) = RestOf(f, 4)
            If Len(selId) > 0 And mSchedId(i) = selId Then mSchedSel = i
            i = i + 1
        Next
    Else
        ReDim mSchedIn(0 To 0)
        ReDim mSchedInN(0 To 0)
        ReDim mSchedNote(0 To 0)
        ReDim mSchedGame(0 To 0)
        ReDim mSchedServer(0 To 0)
        ReDim mSchedConnect(0 To 0)
    End If
    mSchedCount = n
    Set mSchedBuild = Nothing
    ' who's in and the details
    If Not mSchedInfo Is Nothing Then
        For Each v In mSchedInfo
            f = Split(CStr(v), "|")
            k = SchedIndexOf(Fld(f, 1))
            If k < 0 Then
            ElseIf UCase$(f(0)) = "SCHEDGAME" Then
                mSchedGame(k) = Fld(f, 2)
                mSchedServer(k) = Fld(f, 3)
                mSchedConnect(k) = RestOf(f, 4)
            Else
                mSchedInN(k) = Val(Fld(f, 2))
                mSchedIn(k) = Fld(f, 3)
                mSchedNote(k) = RestOf(f, 4)
            End If
        Next
    End If
    Set mSchedInfo = Nothing
    ' reminders for events that are no longer on the schedule
    If Len(mRemind) > 0 Then
        keep = ""
        For i = 0 To mSchedCount - 1
            If IsRemind(i) Then
                If Len(keep) > 0 Then keep = keep & ";"
                keep = keep & RemindKey(i)
            End If
        Next
        If keep <> mRemind Then
            mRemind = keep
            SettingSet "remind", mRemind
        End If
    End If
    UpdateSchedScroll
    DrawSched
    DrawNext
    RefreshWhatsTools
    RefreshOpenForm "frmEventInfo"
    RefreshOpenForm "frmProjector"
End Sub

Private Function VoteSecsLeft() As Long
    Dim s As Long
    s = (mVoteDeadline - GetTickCount()) \ 1000
    If s < 0 Then s = 0
    VoteSecsLeft = s
End Function

Private Function MmSs(ByVal secs As Long) As String
    MmSs = (secs \ 60) & ":" & Format$(secs Mod 60, "00")
End Function

' Once a second: count down the vote, and move "On now" along when the minute changes
Private Sub tmrClock_Timer()
    PerfStart PF_TIMER
    ClockTick
    PerfStop PF_TIMER
End Sub

Private Sub ClockTick()
    Dim hm As String
    ' Windows sometimes knocks over the UDP listener (after a ping to a PC that's gone). Start it again.
    If mFindBroken Then
        mFindBroken = False
        On Error Resume Next
        sckFind.Close
        On Error GoTo 0
        StartListening
    End If
    If Not mSignedIn Then Exit Sub
    ' Who's Here: our pings every 5 seconds, everyone's details every 10
    If mPage = PG_WHO And Me.WindowState <> vbMinimized Then
        mWhoTick = mWhoTick + 1
        If mWhoTick Mod 5 = 0 Then PingPeers
        If mWhoTick Mod 10 = 0 Then RequestWho
        If mPeerDirty Then
            mPeerDirty = False
            If mWhoView = 0 Then DrawPeople
            DrawCard
        End If
    End If
    hm = Format$(Now, "hh:nn")
    If hm <> mLastMinute Then
        mLastMinute = hm
        CheckUpNext
        CheckReminders
        DrawNext
        If mPage = PG_WHATSON Then
            DrawHeader
            DrawSched
        End If
    End If
    If mVoteState = "open" Then
        If mPage = PG_WHATSON Then DrawVote
        If mPage = PG_CHAT Then DrawNext
        If Not mSoonDone And VoteSecsLeft() <= mSoonAt And VoteSecsLeft() > 0 Then
            mSoonDone = True
            If gNotifySoon Then
                Ding
                If mVoteMine = 0 Then
                    ShowToast "soon", IC_CLOCK, SoonTitle(), "You have not voted. Click here to vote.", 10
                Else
                    ShowToast "soon", IC_CLOCK, SoonTitle(), "Your vote: " & mVoteName(mVoteMine - 1) & ". Click here to change it.", 10
                End If
            End If
        End If
    End If
End Sub

' 15, 10 and 5 minutes before the next thing on the schedule, pop up a reminder
Private Sub CheckUpNext()
    Dim nxt As Long, mins As Long, key As String
    If Not gNotifyNext Then Exit Sub
    nxt = CurrentEvent() + 1
    If nxt < 0 Or nxt >= mSchedCount Then Exit Sub
    mins = PartyMinutes(mSchedTime(nxt)) - PartyMinutes(Format$(Now, "hh:nn"))
    If mins <> 15 And mins <> 10 And mins <> 5 Then Exit Sub
    key = mSchedId(nxt) & "@" & mins
    If key = mLastNextKey Then Exit Sub
    mLastNextKey = key
    Ding
    ShowToast "next", mSchedIcon(nxt), "Starting in " & mins & " minutes", mSchedName(nxt) & " at " & AmPm(mSchedTime(nxt)) & ".", 10
End Sub

Private Function SoonTitle() As String
    If mSoonAt >= 60 Then
        SoonTitle = "Voting closes in 1 minute"
    Else
        SoonTitle = "Voting closes in " & mSoonAt & " seconds"
    End If
End Function

' Spots a vote starting or finishing, and pops up a notification about it
Private Sub VoteNews()
    Dim key As String, i As Long, s As String
    key = mVoteQ
    For i = 0 To mVoteCount - 1
        key = key & "|" & mVoteName(i)
    Next
    If mVoteState = "open" And (mPrevVoteState <> "open" Or key <> mVoteKey) Then
        ' a new vote
        mVoteKey = key
        mSoonDone = False
        If VoteSecsLeft() >= 120 Then
            mSoonAt = 60
        Else
            mSoonAt = 30
        End If
        If gNotifyVote Then
            Ding
            ShowToast "vote", IC_STAR, "Vote Started", mVoteQ & " Click here to vote.", 10
        End If
    ElseIf mVoteState = "closed" And mPrevVoteState = "open" Then
        ' it just closed
        If mVoteWinner > 0 Then
            s = mVoteName(mVoteWinner - 1) & " won with " & mVoteVotes(mVoteWinner - 1) & " of " & mVoteTotal & " votes."
        ElseIf mVoteTotal = 0 Then
            s = "No votes were cast."
        Else
            s = "The vote is tied."
        End If
        If gNotifyResult Then ShowToast "result", IC_TROPHY, "Voting Closed", s, 10
    End If
End Sub

' ================================================================ What's On: buttons for everyone

' Everyone's buttons first, then the host's (those are on the Host menu too, so they can go on >>)
Private Sub BuildWhatsTools()
    Dim key As String, sel As Boolean, newSrv As Boolean
    newSrv = ServerHasWhatsOn()
    key = NextGameKey()
    If Len(key) > 0 Then
        AddTool "next", NextGameIcon(key), NextGameText(key), False
        AddTool "-", -1, "", False
    End If
    sel = (mSchedSel >= 0 And mSchedSel < mSchedCount)
    If sel Then
        If newSrv Then AddTool "evin", IC_OK, "Count Me In", IsInEvent(mSchedSel)
        AddTool "remind", IC_BELL, "Remind Me", IsRemind(mSchedSel)
        If newSrv Then AddTool "evinfo", IC_INFO, "Details...", False
        AddTool "evchat", IC_CHAT, "Copy to Chat", False
        AddTool "-", -1, "", False
    End If
    If newSrv Then
        AddTool "suggest", IC_STAR, "Suggest a Game...", False
        AddTool "food", IC_PIZZA, "Food Order...", False
        AddTool "-", -1, "", False
    End If
    AddTool "print", IC_PRINTER, "Print...", False
    If mIsAdmin Then
        AddTool "-", -1, "", False
        AddTool "sadd", IC_PLUS, "Add Event...", False
        AddTool "tpl", IC_CALENDAR, "Templates...", False
        If sel Then AddTool "sdel", IC_ERASER, "Remove Event", False
        AddTool "-", -1, "", False
        If mVoteState = "open" Then
            AddTool "vstop", IC_LOCK, "Close Vote", False
        Else
            AddTool "vstart", IC_STAR, "Start Vote...", False
        End If
    End If
End Sub

' Count Me In, event details, suggestions and the food order need Palms Server 0.4.7
Public Function ServerHasWhatsOn() As Boolean
    ServerHasWhatsOn = (VerNum(mServerVer) >= 407)
End Function

Private Sub ShowEventMenu()
    Dim newSrv As Boolean
    If mSchedSel < 0 Or mSchedSel >= mSchedCount Then Exit Sub
    newSrv = ServerHasWhatsOn()
    mnuEvIn.Visible = newSrv
    mnuEvIn.Checked = IsInEvent(mSchedSel)
    mnuEvRemind.Checked = IsRemind(mSchedSel)
    mnuEvInfo.Visible = newSrv
    mnuEvSep.Visible = mIsAdmin
    mnuEvDel.Visible = mIsAdmin
    PopupMenu mnuEvent
End Sub

Private Sub ShowEventInfo(ByVal i As Long)
    If i < 0 Or i >= mSchedCount Then Exit Sub
    If Not ServerHasWhatsOn() Then Exit Sub
    frmEventInfo.ShowEvent mSchedId(i)
End Sub

' ---- Play/Install Next

' The game to get ready for: the selected event's, the one on now, the vote's (the winner, or the
' one I voted for), or the next thing on the schedule. Only games in the Game Library count.
Private Function NextGameKey() As String
    Dim i As Long, cur As Long, k As String
    If gShowCount = 0 Then Exit Function
    If mSchedSel >= 0 And mSchedSel < mSchedCount Then k = LibFind(mSchedName(mSchedSel))
    cur = CurrentEvent()
    If Len(k) = 0 And cur >= 0 Then k = LibFind(mSchedName(cur))
    If Len(k) = 0 Then
        If mVoteState = "closed" And mVoteWinner > 0 Then
            k = LibFind(mVoteName(mVoteWinner - 1))
        ElseIf mVoteState = "open" And mVoteMine > 0 Then
            k = LibFind(mVoteName(mVoteMine - 1))
        End If
    End If
    If Len(k) = 0 Then
        For i = cur + 1 To mSchedCount - 1
            k = LibFind(mSchedName(i))
            If Len(k) > 0 Then Exit For
        Next
    End If
    NextGameKey = k
End Function

' The Game Library game named in some text (an event, a vote choice). The longest name wins,
' so "Quake III Arena final" finds Quake III Arena rather than Quake.
Private Function LibFind(ByVal text As String) As String
    Dim i As Long, nm As String, best As Long
    If Len(text) = 0 Then Exit Function
    For i = 0 To gShowCount - 1
        nm = LibName(gShowId(i))
        If Len(nm) >= 3 And Len(nm) > best Then
            If InStr(1, text, nm, vbTextCompare) > 0 Then
                LibFind = gShowId(i)
                best = Len(nm)
            End If
        End If
    Next
End Function

Private Function NextGameReady(ByVal key As String) As Boolean
    Select Case LibState(key)
    Case LS_READY, LS_OLD, LS_UPDATE
        NextGameReady = LibInstalled(key)
    End Select
End Function

Private Function NextGameText(ByVal key As String) As String
    Dim nm As String
    nm = LibName(key)
    If Len(nm) > 24 Then nm = Left$(nm, 22) & "..."
    If LibState(key) = LS_BUSY Or LibState(key) = LS_WAIT Then
        NextGameText = "Installing " & nm
    ElseIf NextGameReady(key) Then
        NextGameText = "Play " & nm
    Else
        NextGameText = "Install " & nm
    End If
End Function

Private Function NextGameIcon(ByVal key As String) As Integer
    If NextGameReady(key) And LibState(key) <> LS_BUSY And LibState(key) <> LS_WAIT Then
        NextGameIcon = IC_JOYSTICK
    Else
        NextGameIcon = IC_DOWNLOAD
    End If
End Function

Private Sub NextGameGo()
    Dim key As String
    key = NextGameKey()
    If Len(key) = 0 Then Exit Sub
    If NextGameReady(key) And LibState(key) <> LS_BUSY And LibState(key) <> LS_WAIT Then
        LibPlay key, Me.hWnd, False
        Exit Sub
    End If
    ' not installed yet: over to the Game Library with it selected, and start installing
    mLibSel = key
    SettingSet "lib_sel", key
    GoPage PG_LIBRARY
    If LibState(key) = LS_NONE Then
        If LibAskInstall(key) Then LibRefresh
    End If
    DrawLibPage
End Sub

' ---- Remind Me (kept on this PC, in the registry)

Private Function RemindKey(ByVal i As Long) As String
    RemindKey = mSchedId(i) & "@" & mSchedTime(i)
End Function

Public Function IsRemind(ByVal i As Long) As Boolean
    If i < 0 Or i >= mSchedCount Or Len(mRemind) = 0 Then Exit Function
    IsRemind = (InStr(1, ";" & mRemind & ";", ";" & RemindKey(i) & ";", vbBinaryCompare) > 0)
End Function

Public Sub ToggleRemind(ByVal i As Long)
    Dim k As String
    If i < 0 Or i >= mSchedCount Then Exit Sub
    k = RemindKey(i)
    If IsRemind(i) Then
        mRemind = Mid$(Replace(";" & mRemind & ";", ";" & k & ";", ";"), 2)
        If Right$(mRemind, 1) = ";" Then mRemind = Left$(mRemind, Len(mRemind) - 1)
        Note "Reminder removed: " & mSchedName(i) & "."
    Else
        If Len(mRemind) > 0 Then mRemind = mRemind & ";"
        mRemind = mRemind & k
        Note "Reminder set: " & mSchedName(i) & ", " & REMIND_MINS & " minutes before " & AmPm(mSchedTime(i)) & "."
    End If
    SettingSet "remind", mRemind
    DrawSched
    RefreshWhatsTools
    RefreshOpenForm "frmEventInfo"
End Sub

' Once a minute: pop up 5 minutes before (or straight away, if the reminder was set later than that)
Private Sub CheckReminders()
    Dim i As Long, mins As Long, k As String, t As String
    If Len(mRemind) = 0 Then Exit Sub
    For i = 0 To mSchedCount - 1
        If IsRemind(i) Then
            mins = PartyMinutes(mSchedTime(i)) - PartyMinutes(Format$(Now, "hh:nn"))
            k = RemindKey(i)
            If mins >= 0 And mins <= REMIND_MINS And InStr(1, ";" & mRemindDone & ";", ";" & k & ";", vbBinaryCompare) = 0 Then
                mRemindDone = mRemindDone & ";" & k
                If mins = 0 Then
                    t = "Starting now"
                ElseIf mins = 1 Then
                    t = "Starting in 1 minute"
                Else
                    t = "Starting in " & mins & " minutes"
                End If
                PlayChime "ding"
                ShowToast "remind", mSchedIcon(i), t, mSchedName(i) & " at " & AmPm(mSchedTime(i)) & ".", 15
            End If
        End If
    Next
End Sub

' ---- Count Me In

Public Function IsInEvent(ByVal i As Long) As Boolean
    If i < 0 Or i >= mSchedCount Or Len(mMyNick) = 0 Then Exit Function
    IsInEvent = (InStr(1, "," & LCase$(mSchedIn(i)) & ",", "," & LCase$(mMyNick) & ",", vbBinaryCompare) > 0)
End Function

Public Sub ToggleIn(ByVal i As Long)
    If i < 0 Or i >= mSchedCount Then Exit Sub
    If Not ServerHasWhatsOn() Then Exit Sub
    If IsInEvent(i) Then
        SendLine "SCHEDOUT|" & mSchedId(i)
        Note "Counted out: " & mSchedName(i) & "."
    Else
        SendLine "SCHEDIN|" & mSchedId(i)
        Note "Counted in: " & mSchedName(i) & "."
    End If
End Sub

' ---- Copy to Chat: the event, typed into the chat box ready to send

Private Sub EventToChat(ByVal i As Long)
    Dim s As String, cur As Long
    If i < 0 Or i >= mSchedCount Then Exit Sub
    cur = CurrentEvent()
    If i = cur Then
        s = "On now: " & mSchedName(i) & " (started " & AmPm(mSchedTime(i)) & ")"
    ElseIf i = cur + 1 Then
        s = "Up next: " & mSchedName(i) & " at " & AmPm(mSchedTime(i))
    Else
        s = AmPm(mSchedTime(i)) & ": " & mSchedName(i)
    End If
    If mSchedInN(i) = 1 Then
        s = s & " - 1 player in"
    ElseIf mSchedInN(i) > 1 Then
        s = s & " - " & mSchedInN(i) & " players in"
    End If
    GoPage PG_CHAT
    txtSay.Text = Left$(s, txtSay.MaxLength)
    txtSay.SelStart = Len(txtSay.Text)
    On Error Resume Next
    txtSay.SetFocus
End Sub

' ---- File > Print Schedule

Private Sub PrintSchedule()
    Dim i As Long, k As Long, parts() As String, dev As String, x2 As Single, maxW As Single
    If Not mSignedIn Then Exit Sub
    If mSchedCount = 0 Then
        MsgBox "There are no events on today's schedule.", vbInformation, "Print Schedule"
        Exit Sub
    End If
    On Error GoTo NoPrinter
    dev = Printer.DeviceName
    On Error GoTo 0
    If MsgBox("Print today's schedule on " & dev & "?", vbOKCancel + vbQuestion, "Print Schedule") <> vbOK Then Exit Sub
    On Error GoTo Failed
    Screen.MousePointer = vbHourglass
    Printer.ScaleMode = vbTwips
    Printer.FontName = "Arial"
    x2 = 2880
    maxW = Printer.ScaleWidth - x2 - 1080
    Printer.CurrentY = 1080
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.CurrentX = 1080
    Printer.Print "Today's Schedule"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = 1080
    Printer.Print mServerName & " - " & Format$(Now, "dddd d mmmm yyyy")
    Printer.Print
    For i = 0 To mSchedCount - 1
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = 1080
        Printer.Print AmPm(mSchedTime(i));
        Printer.FontBold = False
        Printer.CurrentX = x2
        Printer.Print mSchedName(i)
        Printer.FontSize = 9
        If Len(mSchedNote(i)) > 0 Then
            parts = Split(mSchedNote(i), "~")
            For k = 0 To UBound(parts)
                If Len(Trim$(parts(k))) > 0 Then PrintWrapped x2, Trim$(parts(k)), maxW
            Next
        End If
        If mSchedInN(i) > 0 Then PrintWrapped x2, "Players in (" & mSchedInN(i) & "): " & Replace(mSchedIn(i), ",", ", "), maxW
        Printer.Print
    Next
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    Note "Schedule sent to " & dev & "."
    Exit Sub
NoPrinter:
    MsgBox "No printer is installed on this PC.", vbExclamation, "Print Schedule"
    Exit Sub
Failed:
    Screen.MousePointer = vbDefault
    MsgBox "Unable to print (" & Err.Description & ").", vbExclamation, "Print Schedule"
    On Error Resume Next
    Printer.KillDoc
End Sub

' One line of text on the printout, wrapped at maxW
Private Sub PrintWrapped(ByVal x As Single, ByVal s As String, ByVal maxW As Single)
    Dim words() As String, ln As String, k As Long
    words = Split(s, " ")
    For k = 0 To UBound(words)
        If Len(ln) = 0 Then
            ln = words(k)
        ElseIf Printer.TextWidth(ln & " " & words(k)) > maxW Then
            Printer.CurrentX = x
            Printer.Print ln
            ln = words(k)
        Else
            ln = ln & " " & words(k)
        End If
    Next
    If Len(ln) > 0 Then
        Printer.CurrentX = x
        Printer.Print ln
    End If
End Sub

' ---- the food order opening pops up a notification, like an announcement from the host

Private Sub FoodNews(ByVal st As String, ByVal title As String)
    If mFoodHeard And mFoodState <> "open" And st = "open" Then
        If gNotifyPA Then
            Ding
            ShowToast "food", IC_PIZZA, "Food Order Open", title & ". Click here to order.", 10
        End If
    End If
    mFoodHeard = True
End Sub

' A What's On window that's open shows the latest
Private Sub RefreshOpenForm(ByVal nm As String)
    If Not IsFormLoaded(nm) Then Exit Sub
    Select Case nm
    Case "frmSuggest"
        frmSuggest.RefreshList
    Case "frmFood"
        frmFood.RefreshFood
    Case "frmEventInfo"
        frmEventInfo.RefreshInfo
    Case "frmVote"
        frmVote.RefreshSugg
    Case "frmProjector"
        frmProjector.DataChanged
    End Select
End Sub

' ---- for the Event Details, Suggest a Game, Food Order, Add Event and Start Vote windows

Public Function HostMode() As Boolean
    HostMode = mIsAdmin
End Function

Public Function SchedIndexOf(ByVal id As String) As Long
    Dim i As Long
    SchedIndexOf = -1
    For i = 0 To mSchedCount - 1
        If mSchedId(i) = id Then
            SchedIndexOf = i
            Exit Function
        End If
    Next
End Function

' what: id, time, when (like 7:30 PM), name, note (~ between lines) or in (nicknames, separated by commas)
Public Function SchedText(ByVal i As Long, ByVal what As String) As String
    If i < 0 Or i >= mSchedCount Then Exit Function
    Select Case what
    Case "id": SchedText = mSchedId(i)
    Case "time": SchedText = mSchedTime(i)
    Case "when": SchedText = AmPm(mSchedTime(i))
    Case "name": SchedText = mSchedName(i)
    Case "note": SchedText = mSchedNote(i)
    Case "in": SchedText = mSchedIn(i)
    Case "game": SchedText = mSchedGame(i)
    Case "server": SchedText = mSchedServer(i)
    Case "connect": SchedText = mSchedConnect(i)
    End Select
End Function

Public Function SchedIconOf(ByVal i As Long) As Integer
    If i >= 0 And i < mSchedCount Then SchedIconOf = mSchedIcon(i)
End Function

' Add Event: the details are sent once the server says the new event's ID (SCHEDADDED)
Public Sub SetPendingNote(ByVal s As String)
    mPendingNote = s
End Sub

' Add Event: game|server|how to connect, sent once the server says the new event's ID
Public Sub SetPendingGame(ByVal s As String)
    mPendingGame = s
End Sub

' ---- for the projector

Public Function SchedCount() As Long
    SchedCount = mSchedCount
End Function

' The event on now (the last one that has started), or -1
Public Function SchedNowIndex() As Long
    Dim i As Long, nowM As Long
    SchedNowIndex = -1
    nowM = PartyMinutes(Format$(Time, "hh:nn"))
    For i = 0 To mSchedCount - 1
        If PartyMinutes(mSchedTime(i)) <= nowM Then
            If SchedNowIndex < 0 Then
                SchedNowIndex = i
            ElseIf PartyMinutes(mSchedTime(i)) >= PartyMinutes(mSchedTime(SchedNowIndex)) Then
                SchedNowIndex = i
            End If
        End If
    Next
End Function

' The next event to start, or -1
Public Function SchedNextIndex() As Long
    Dim i As Long, nowM As Long
    SchedNextIndex = -1
    nowM = PartyMinutes(Format$(Time, "hh:nn"))
    For i = 0 To mSchedCount - 1
        If PartyMinutes(mSchedTime(i)) > nowM Then
            If SchedNextIndex < 0 Then
                SchedNextIndex = i
            ElseIf PartyMinutes(mSchedTime(i)) < PartyMinutes(mSchedTime(SchedNextIndex)) Then
                SchedNextIndex = i
            End If
        End If
    Next
End Function

Public Function ServerHasProjector() As Boolean
    ServerHasProjector = (VerNum(mServerVer) >= 600)
End Function

' Seconds per slide, and the slides (now, sched, cup, vote), as the host set them
Public Function ProjSeconds() As Long
    ProjSeconds = mProjSecs
    If ProjSeconds < 5 Then ProjSeconds = 15
End Function

Public Function ProjSlides() As String
    If Len(mProjPort) = 0 Then
        ProjSlides = "now,sched,cup,vote"
    Else
        ProjSlides = mProjSlides
    End If
End Function

' The projector web page, like http://10.0.200.71:8080/
Public Function ProjUrl() As String
    If Len(mProjPort) > 0 And Len(mServerHost) > 0 Then ProjUrl = "http://" & mServerHost & ":" & mProjPort & "/"
End Function

Public Function PalmsServerName() As String
    PalmsServerName = mServerName
End Function

' The vote: what is state (none, open or closed), question, count, total, winner or left (seconds)
Public Function VoteText(ByVal what As String) As String
    Dim s As Long
    Select Case what
    Case "state": VoteText = mVoteState
    Case "question": VoteText = mVoteQ
    Case "count": VoteText = CStr(mVoteCount)
    Case "total": VoteText = CStr(mVoteTotal)
    Case "winner": VoteText = CStr(mVoteWinner)
    Case "left"
        s = (mVoteDeadline - GetTickCount()) \ 1000
        If s < 0 Then s = 0
        VoteText = CStr(s)
    End Select
End Function

Public Function VoteChoice(ByVal i As Long) As String
    If i >= 0 And i < mVoteCount And i <= 5 Then VoteChoice = mVoteName(i)
End Function

Public Function VoteVotes(ByVal i As Long) As Long
    If i >= 0 And i < mVoteCount And i <= 5 Then VoteVotes = mVoteVotes(i)
End Function

Public Function SuggCount() As Long
    SuggCount = mSuggCount
End Function

Public Function SuggName(ByVal i As Long) As String
    If i >= 0 And i < mSuggCount Then SuggName = mSuggName(i)
End Function

Public Function SuggBy(ByVal i As Long) As String
    If i >= 0 And i < mSuggCount Then SuggBy = mSuggBy(i)
End Function

Public Function SuggVotes(ByVal i As Long) As Long
    If i >= 0 And i < mSuggCount Then SuggVotes = mSuggN(i)
End Function

Public Function SuggMine(ByVal i As Long) As Boolean
    If i < 0 Or i >= mSuggCount Or Len(mMyNick) = 0 Then Exit Function
    SuggMine = (InStr(1, "," & LCase$(mSuggBy(i)) & ",", "," & LCase$(mMyNick) & ",", vbBinaryCompare) > 0)
End Function

Public Function SuggMineCount() As Long
    Dim i As Long
    For i = 0 To mSuggCount - 1
        If SuggMine(i) Then SuggMineCount = SuggMineCount + 1
    Next
End Function

Public Function FoodState() As String
    FoodState = mFoodState
End Function

Public Function FoodTitle() As String
    FoodTitle = mFoodTitle
End Function

Public Function FoodCount() As Long
    FoodCount = mFoodCount
End Function

Public Function FoodNick(ByVal i As Long) As String
    If i >= 0 And i < mFoodCount Then FoodNick = mFoodNick(i)
End Function

Public Function FoodTime(ByVal i As Long) As String
    If i >= 0 And i < mFoodCount Then FoodTime = mFoodTime(i)
End Function

Public Function FoodText(ByVal i As Long) As String
    If i >= 0 And i < mFoodCount Then FoodText = mFoodText(i)
End Function

Public Function MyFood() As String
    Dim i As Long
    For i = 0 To mFoodCount - 1
        If StrComp(mFoodNick(i), mMyNick, vbTextCompare) = 0 Then
            MyFood = mFoodText(i)
            Exit Function
        End If
    Next
End Function

' ================================================================ Tournaments
' Single elimination. modCups keeps the tournaments the server sends and draws
' the bracket. This part lays out the page: a tab for each tournament, the bracket (or the sign-ups)
' with scroll bars, and a bar along the bottom about the match picked, or where you stand.

Public Function ServerHasCups() As Boolean
    ServerHasCups = (VerNum(mServerVer) >= 500)
End Function

' The tournament on show: where modCups keeps it, or -1
Private Function CupK() As Long
    CupK = CupIndex(mCupSel)
End Function

Private Sub CupGeometry(w As Long, h As Long, panelBottom As Long, stripY As Long)
    w = picCupsPage.ScaleWidth
    h = picCupsPage.ScaleHeight
    stripY = h - 10 - 46
    panelBottom = stripY - 10
    If panelBottom < TAB_BOTTOM + 80 Then panelBottom = TAB_BOTTOM + 80
End Sub

Private Sub LayoutCups()
    Dim w As Long, h As Long, pb As Long, sy As Long
    CupGeometry w, h, pb, sy
    picCup.Move 16, TAB_BOTTOM + 6, w - 32, pb - TAB_BOTTOM - 12
    cmdCupGo.Move w - 10 - 8 - 140, sy + 11, 140, 25
    UpdateCupScroll
    DrawCupPage
End Sub

' How far the bracket is scrolled, in pixels
Private Function CupSX() As Long
    If hsbCup.Visible Then CupSX = hsbCup.Value
End Function

Private Function CupSY() As Long
    If vsbCup.Visible Then CupSY = vsbCup.Value
End Function

' The scroll bars: only the ones needed, sized to what's on show
Private Sub UpdateCupScroll()
    Dim k As Long, cw As Long, ch As Long, pw As Long, ph As Long, needV As Boolean, needH As Boolean, mx As Long
    k = CupK()
    pw = picCup.ScaleWidth
    ph = picCup.ScaleHeight
    If k >= 0 Then
        cw = CupViewW(k, pw)
        ch = CupViewH(picCup, k, pw)
        If ch > ph Then
            needV = True
            pw = pw - 16
            cw = CupViewW(k, pw)
            ch = CupViewH(picCup, k, pw)
        End If
        If cw > pw Then
            needH = True
            ph = ph - 16
            If Not needV And ch > ph Then
                needV = True
                pw = pw - 16
            End If
        End If
    End If
    mIgnoreScroll = True
    If needV Then
        vsbCup.Move picCup.ScaleWidth - 16, 0, 16, ph
        mx = ch - ph
        If vsbCup.Value > mx Then vsbCup.Value = mx
        vsbCup.Max = mx
        vsbCup.SmallChange = 20
        If ph > 40 Then vsbCup.LargeChange = ph - 20 Else vsbCup.LargeChange = 20
    Else
        vsbCup.Value = 0
    End If
    If needH Then
        hsbCup.Move 0, picCup.ScaleHeight - 16, pw, 16
        mx = cw - pw
        If hsbCup.Value > mx Then hsbCup.Value = mx
        hsbCup.Max = mx
        hsbCup.SmallChange = 20
        If pw > 40 Then hsbCup.LargeChange = pw - 20 Else hsbCup.LargeChange = 20
    Else
        hsbCup.Value = 0
    End If
    If vsbCup.Visible <> needV Then vsbCup.Visible = needV
    If hsbCup.Visible <> needH Then hsbCup.Visible = needH
    mIgnoreScroll = False
End Sub

Private Sub SetCupScroll(ByVal x As Long, ByVal y As Long)
    If hsbCup.Visible Then
        If x > hsbCup.Max Then x = hsbCup.Max
        If x < 0 Then x = 0
        hsbCup.Value = x
    End If
    If vsbCup.Visible Then
        If y > vsbCup.Max Then y = vsbCup.Max
        If y < 0 Then y = 0
        vsbCup.Value = y
    End If
End Sub

' Scrolls a match into sight
Private Sub CupScrollTo(ByVal m As Long)
    Dim x As Long, y As Long, w As Long, h As Long, pw As Long, ph As Long, sx As Long, sy As Long
    If m < 0 Then Exit Sub
    CupMatchSpot m \ 100, m Mod 100, x, y, w, h
    pw = picCup.ScaleWidth
    ph = picCup.ScaleHeight
    If vsbCup.Visible Then pw = pw - 16
    If hsbCup.Visible Then ph = ph - 16
    sx = CupSX()
    sy = CupSY()
    If x - 10 < sx Then sx = x - 10
    If x + w + 10 > sx + pw Then sx = x + w + 10 - pw
    If y - 30 < sy Then sy = y - 30
    If y + h + 10 > sy + ph Then sy = y + h + 10 - ph
    SetCupScroll sx, sy
End Sub

Private Sub DrawCupPage()
    Want DR_CUPPAGE
End Sub

Private Sub DrawCupView()
    Want DR_CUP
End Sub

Private Sub DrawCupPageNow()
    Dim w As Long, h As Long, pb As Long, sy As Long, k As Long, s As String, ic As Integer
    Dim l1 As String, l2 As String, col2 As Long, tw As Long
    CupGeometry w, h, pb, sy
    picCupsPage.Cls
    k = CupK()
    If k < 0 Then
        ' nothing on: an empty panel, and what to do about it
        If picCup.Visible Then picCup.Visible = False
        Edge picCupsPage, 10, TAB_BOTTOM, w - 20, pb - TAB_BOTTOM, vb3DHighlight, vb3DDKShadow
        Edge picCupsPage, 11, TAB_BOTTOM + 1, w - 22, pb - TAB_BOTTOM - 2, vbButtonFace, vb3DShadow
        DrawIcon32 picCupsPage, IC_TROPHY, (w - 32) \ 2, TAB_BOTTOM + 40
        s = "No tournaments."
        TextAt picCupsPage, (w - picCupsPage.TextWidth(s)) \ 2, TAB_BOTTOM + 82, s, C_SUB
        If mIsAdmin Then
            s = "To start one, click New Tournament on the toolbar."
        Else
            s = "Tournaments started by the host are shown here."
        End If
        TextAt picCupsPage, (w - picCupsPage.TextWidth(s)) \ 2, TAB_BOTTOM + 100, s, C_GREY
        TipBox picCupsPage, 10, sy, w - 20, 46
        DrawIcon16 picCupsPage, IC_INFO, 18, sy + 7
        TextAt picCupsPage, 40, sy + 8, "Sign up for a tournament here when the host opens one.", vbBlack, True
        TextAt picCupsPage, 40, sy + 25, "When your match is ready, a notification appears.", C_SUB
        If cmdCupGo.Visible Then cmdCupGo.Visible = False
        Exit Sub
    End If
    If Not picCup.Visible Then picCup.Visible = True
    DrawCupTabs w - 20, pb
    ' the bar along the bottom
    CupStripText k, ic, l1, l2, col2
    TipBox picCupsPage, 10, sy, w - 20, 46
    DrawIcon16 picCupsPage, ic, 18, sy + 7
    UpdateCupButton
    tw = w - 20 - 40 - 8
    If cmdCupGo.Visible Then tw = tw - 150
    picCupsPage.FontBold = True
    l1 = FitText(picCupsPage, l1, tw)
    picCupsPage.FontBold = False
    TextAt picCupsPage, 40, sy + 8, l1, vbBlack, True
    TextAt picCupsPage, 40, sy + 25, FitText(picCupsPage, l2, tw), col2
    DrawCupNow
End Sub

' The two lines along the bottom: the match picked, or where you stand
Private Sub CupStripText(ByVal k As Long, ic As Integer, l1 As String, l2 As String, col2 As Long)
    Dim s As String, parts() As String, n As Long, r As Long, i As Long
    col2 = C_SUB
    ic = IC_INFO
    If mCupMatch >= 0 Then
        r = mCupMatch \ 100
        i = mCupMatch Mod 100
        s = CupMatchText(k, r, i, mMyNick, mIsAdmin)
    End If
    If Len(s) > 0 Then
        parts = Split(s & vbLf, vbLf)
        l1 = parts(0)
        l2 = parts(1)
        If Len(CupM(k, r, i, "repby")) > 0 And Len(CupM(k, r, i, "win")) = 0 Then ic = IC_WARN
    Else
        l1 = CupStatusText(k, mMyNick)
        l2 = CupInfoText(k)
        If Len(gCupChamp(k)) > 0 Then
            ic = IC_TROPHY
        ElseIf CupMyTurn(k, mMyNick) Then
            ic = IC_BELL
        End If
        ' the host: what to do next
        If mIsAdmin And gCupState(k) = "signup" Then
            l2 = "When everyone has signed up, click Draw Bracket on the toolbar."
        ElseIf mIsAdmin And gCupState(k) = "drawn" Then
            l2 = "To begin, click Start on the toolbar. To change the pairs first, click Swap Players."
        End If
        If mIsAdmin And Len(gCupChamp(k)) = 0 Then
            n = CupPending(k)
            If n = 1 Then
                ic = IC_WARN
                l2 = "1 result is waiting for you to confirm. It is marked with a warning sign."
            ElseIf n > 1 Then
                ic = IC_WARN
                l2 = n & " results are waiting for you to confirm. They are marked with a warning sign."
            End If
        End If
    End If
    If Len(mCupError) > 0 Then
        l2 = mCupError
        col2 = C_RED
        ic = IC_WARN
    End If
End Sub

Private Sub DrawCupNow()
    Dim k As Long
    k = CupK()
    If k < 0 Then
        picCup.Cls
        Exit Sub
    End If
    DrawCup picCup, k, CupSX(), CupSY(), mCupMatch, mMyNick
    ' the corner between the two scroll bars
    If vsbCup.Visible And hsbCup.Visible Then Box picCup, picCup.ScaleWidth - 16, picCup.ScaleHeight - 16, 16, 16, vbButtonFace
End Sub

' ---- the tabs, one for each tournament. They share the width when they don't all fit.

Private Function CupTabIcon(ByVal k As Long) As Integer
    If mIsAdmin And CupPending(k) > 0 Then
        CupTabIcon = IC_WARN
    ElseIf CupMyTurn(k, mMyNick) Then
        CupTabIcon = IC_BELL
    Else
        Select Case gCupState(k)
        Case "signup": CupTabIcon = IC_PEOPLE
        Case "drawn": CupTabIcon = IC_DICE
        Case "running": CupTabIcon = IC_FLAG
        Case Else: CupTabIcon = IC_TROPHY
        End Select
    End If
End Function

Private Sub DrawCupTabs(ByVal lw As Long, ByVal panelBottom As Long)
    Dim pass As Integer, k As Long, x As Long, w As Long, y As Long, h As Long, cap As Long, tw As Long, sel As Long, s As String
    sel = CupK()
    Edge picCupsPage, 10, TAB_BOTTOM, lw, panelBottom - TAB_BOTTOM, vb3DHighlight, vb3DDKShadow
    Edge picCupsPage, 11, TAB_BOTTOM + 1, lw - 2, panelBottom - TAB_BOTTOM - 2, vbButtonFace, vb3DShadow
    cap = (lw - 8) \ gCupCount - 1
    If cap > 200 Then cap = 200
    picCupsPage.FontBold = True
    x = 12
    For k = 0 To gCupCount - 1
        tw = picCupsPage.TextWidth(gCupName(k))
        w = 6 + 16 + 5 + tw + 10
        If w > cap Then w = cap
        mCupTabX(k) = x
        mCupTabW(k) = w
        x = x + w + 1
    Next
    picCupsPage.FontBold = False
    ' the other tabs first, then the one picked, so it sits on top
    For pass = 0 To 1
        For k = 0 To gCupCount - 1
            If (k = sel) = (pass = 1) Then
                x = mCupTabX(k)
                w = mCupTabW(k)
                If k = sel Then
                    x = x - 2
                    w = w + 4
                    y = TAB_BOTTOM - 22
                    h = 23
                Else
                    y = TAB_BOTTOM - 20
                    h = 20
                End If
                Box picCupsPage, x, y, w, h, vbButtonFace
                Box picCupsPage, x, y + 2, 1, h - 2, vb3DHighlight
                Box picCupsPage, x + 1, y + 1, 1, 1, vb3DHighlight
                Box picCupsPage, x + 2, y, w - 4, 1, vb3DHighlight
                Box picCupsPage, x + w - 2, y + 1, 1, 1, vb3DDKShadow
                Box picCupsPage, x + w - 1, y + 2, 1, h - 2, vb3DDKShadow
                Box picCupsPage, x + w - 2, y + 2, 1, h - 2, vb3DShadow
                DrawIcon16 picCupsPage, CupTabIcon(k), x + 6, y + 3
                If k = sel Then picCupsPage.FontBold = True
                s = FitText(picCupsPage, gCupName(k), w - 27 - 6)
                picCupsPage.FontBold = False
                TextAt picCupsPage, x + 27, y + 5, s, vbButtonText, (k = sel)
            End If
        Next
    Next
End Sub

Private Function CupTabAt(ByVal px As Single, ByVal py As Single) As Long
    Dim k As Long
    CupTabAt = -1
    If py < TAB_BOTTOM - 22 Or py >= TAB_BOTTOM Then Exit Function
    For k = 0 To gCupCount - 1
        If px >= mCupTabX(k) And px < mCupTabX(k) + mCupTabW(k) Then CupTabAt = k
    Next
End Function

Private Sub picCupsPage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k As Long
    If Button <> vbLeftButton Then Exit Sub
    k = CupTabAt(X, Y)
    If k >= 0 Then
        If gCupId(k) <> mCupSel Then SelectCup gCupId(k)
    End If
End Sub

' Show a tournament: no match picked, scrolled to your own match if you have one
Private Sub SelectCup(ByVal id As String)
    Dim k As Long
    mCupSel = id
    mCupMatch = -1
    mCupError = ""
    mIgnoreScroll = True
    hsbCup.Value = 0
    vsbCup.Value = 0
    mIgnoreScroll = False
    UpdateCupScroll
    k = CupK()
    If k >= 0 Then CupScrollTo CupMyMatch(k, mMyNick)
    DrawCupPage
    RefreshCupTools
    On Error Resume Next
    If mPage = PG_CUPS And picCup.Visible Then picCup.SetFocus
End Sub

' ---- the bracket: click a match to pick it, double-click to report or enter its result

Private Sub picCup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k As Long, m As Long
    k = CupK()
    If k < 0 Then Exit Sub
    m = CupHit(k, CLng(X) + CupSX(), CLng(Y) + CupSY())
    If m <> mCupMatch Or Len(mCupError) > 0 Then
        mCupMatch = m
        mCupError = ""
        DrawCupPage
        RefreshCupTools
    End If
End Sub

Private Sub picCup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim k As Long
    k = CupK()
    If k < 0 Then Exit Sub
    If CupHit(k, CLng(X) + CupSX(), CLng(Y) + CupSY()) >= 0 Then HandCursor
End Sub

Private Sub picCup_DblClick()
    If mCupMatch >= 0 Then CupResultDialog mCupMatch
End Sub

Private Sub picCup_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp: SetCupScroll CupSX(), CupSY() - 20
    Case vbKeyDown: SetCupScroll CupSX(), CupSY() + 20
    Case vbKeyLeft: SetCupScroll CupSX() - 20, CupSY()
    Case vbKeyRight: SetCupScroll CupSX() + 20, CupSY()
    Case vbKeyPageUp: SetCupScroll CupSX(), CupSY() - vsbCup.LargeChange
    Case vbKeyPageDown: SetCupScroll CupSX(), CupSY() + vsbCup.LargeChange
    Case vbKeyHome: SetCupScroll 0, 0
    Case vbKeyEnd: SetCupScroll hsbCup.Max, vsbCup.Max
    Case vbKeyReturn: If mCupMatch >= 0 Then CupResultDialog mCupMatch
    End Select
End Sub

Private Sub vsbCup_Change()
    If Not mIgnoreScroll Then DrawCupView
End Sub

Private Sub vsbCup_Scroll()
    DrawCupView
End Sub

Private Sub hsbCup_Change()
    If Not mIgnoreScroll Then DrawCupView
End Sub

Private Sub hsbCup_Scroll()
    DrawCupView
End Sub

' ---- the buttons

' This PC's player can report the result of their match now
Private Function CupCanReport(ByVal k As Long) As Boolean
    Dim m As Long
    If gCupState(k) <> "running" Then Exit Function
    m = CupMyMatch(k, mMyNick)
    If m < 0 Then Exit Function
    CupCanReport = CupReady(k, m \ 100, m Mod 100)
End Function

' The host can enter a result for the match picked: both players known
Private Function CupCanEnter(ByVal k As Long) As Boolean
    Dim a As String, b As String
    If mCupMatch < 0 Or Not mIsAdmin Then Exit Function
    If gCupState(k) <> "running" And gCupState(k) <> "done" Then Exit Function
    a = CupM(k, mCupMatch \ 100, mCupMatch Mod 100, "a")
    b = CupM(k, mCupMatch \ 100, mCupMatch Mod 100, "b")
    CupCanEnter = (Len(a) > 0 And a <> "-" And Len(b) > 0 And b <> "-")
End Function

' A result reported for the match picked, waiting for the host
Private Function CupPickedPending(ByVal k As Long) As Boolean
    If mCupMatch < 0 Then Exit Function
    CupPickedPending = Len(CupM(k, mCupMatch \ 100, mCupMatch Mod 100, "repby")) > 0 And Len(CupM(k, mCupMatch \ 100, mCupMatch Mod 100, "win")) = 0
End Function

' Everyone's buttons first, then the host's (they're on the Host menu too, so they can go on >>)
Private Sub BuildCupTools()
    Dim k As Long, st As String
    k = CupK()
    If k >= 0 Then
        st = gCupState(k)
        If st = "signup" Then
            If CupIsIn(k, mMyNick) Then
                AddTool "cupleave", IC_ERASER, "Withdraw", False
            Else
                AddTool "cupjoin", IC_PLUS, "Sign Up", False
            End If
        End If
        If CupCanReport(k) Then AddTool "cupreport", IC_FLAG, "Report Result...", False
        If CupMyMatch(k, mMyNick) >= 0 Then AddTool "cupmine", IC_JOYSTICK, "My Match", False
        AddTool "cupchat", IC_CHAT, "Copy to Chat", False
    End If
    If mIsAdmin Then
        If k >= 0 Then AddTool "-", -1, "", False
        AddTool "cupnew", IC_PLUS, "New Tournament...", False
        If k >= 0 Then
            Select Case st
            Case "signup"
                AddTool "cupplayers", IC_PEOPLE, "Players...", False
                AddTool "cupdraw", IC_DICE, "Draw Bracket", False
            Case "drawn"
                AddTool "cupstart", IC_FLAG, "Start", False
                AddTool "cupplayers", IC_PEOPLE, "Swap Players...", False
                AddTool "cupredraw", IC_DICE, "Redraw", False
                AddTool "cupopen", IC_REFRESH, "Reopen Sign-ups", False
            Case Else
                If CupPickedPending(k) Then
                    AddTool "cupresult", IC_OK, "Confirm Result...", False
                ElseIf CupCanEnter(k) Then
                    AddTool "cupresult", IC_PENCIL, "Enter Result...", False
                End If
            End Select
            AddTool "cupedit", IC_PENCIL, "Edit...", False
            AddTool "cupdel", IC_ERASER, "Delete", False
        End If
    End If
End Sub

Private Sub RefreshCupTools()
    If mPage = PG_CUPS Then
        BuildTools
        DrawTools
    End If
End Sub

' The button on the bar along the bottom: the one thing most likely wanted next
Private Sub UpdateCupButton()
    Dim k As Long, key As String, txt As String
    k = CupK()
    If k >= 0 Then
        If gCupState(k) = "signup" Then
            If CupIsIn(k, mMyNick) Then
                key = "cupleave"
                txt = "&Withdraw"
            ElseIf gCupMax(k) = 0 Or CupPlayerCount(k) < gCupMax(k) Then
                key = "cupjoin"
                txt = "&Sign Up"
            End If
        ElseIf mIsAdmin And CupPickedPending(k) Then
            key = "cupresult"
            txt = "&Confirm Result..."
        ElseIf CupCanReport(k) Then
            key = "cupreport"
            txt = "&Report Result..."
        End If
    End If
    mCupGoKey = key
    If Len(key) > 0 And cmdCupGo.Caption <> txt Then cmdCupGo.Caption = txt
    If cmdCupGo.Visible <> (Len(key) > 0) Then cmdCupGo.Visible = (Len(key) > 0)
End Sub

Private Sub cmdCupGo_Click()
    If Len(mCupGoKey) > 0 Then CupAction mCupGoKey
End Sub

Private Sub CupAction(ByVal key As String)
    Dim k As Long, m As Long, n As Long
    k = CupK()
    If k < 0 And key <> "cupnew" Then Exit Sub
    Select Case key
    Case "cupjoin"
        SendLine "CUPJOIN|" & gCupId(k)
        Note "Signed up for " & gCupName(k) & "."
    Case "cupleave"
        SendLine "CUPLEAVE|" & gCupId(k)
        Note "Withdrawn from " & gCupName(k) & "."
    Case "cupreport"
        m = CupMyMatch(k, mMyNick)
        If m >= 0 Then
            mCupMatch = m
            CupScrollTo m
            DrawCupPage
            CupResultDialog m
        End If
    Case "cupresult"
        If mCupMatch >= 0 Then CupResultDialog mCupMatch
    Case "cupmine"
        m = CupMyMatch(k, mMyNick)
        If m >= 0 Then
            mCupMatch = m
            CupScrollTo m
            DrawCupPage
            RefreshCupTools
        End If
    Case "cupchat"
        CupToChat k
    Case "cupnew"
        If ServerHasCups() Then frmCupEdit.EditCup ""
    Case "cupedit"
        frmCupEdit.EditCup gCupId(k)
    Case "cupplayers"
        frmCupPlayers.ShowCup gCupId(k)
    Case "cupdraw"
        n = CupPlayerCount(k)
        If n < 2 Then
            MsgBox "At least 2 players must sign up before the bracket can be drawn.", vbExclamation, "Draw Bracket"
        ElseIf MsgBox("Close sign-ups for " & gCupName(k) & " and draw the bracket? " & n & " players have signed up." & vbCrLf & vbCrLf & "The players are placed at random. The tournament starts when you click Start.", vbQuestion + vbOKCancel, "Draw Bracket") = vbOK Then
            SendLine "CUPDRAW|" & gCupId(k)
        End If
    Case "cupredraw"
        If MsgBox("Draw the bracket for " & gCupName(k) & " again? The players are placed at random.", vbQuestion + vbOKCancel, "Redraw") = vbOK Then SendLine "CUPDRAW|" & gCupId(k)
    Case "cupstart"
        If MsgBox("Start " & gCupName(k) & "? Each player is told about their first match.", vbQuestion + vbOKCancel, "Start Tournament") = vbOK Then SendLine "CUPSTART|" & gCupId(k)
    Case "cupopen"
        If MsgBox("Reopen sign-ups for " & gCupName(k) & "? The bracket is cleared.", vbQuestion + vbOKCancel, "Reopen Sign-ups") = vbOK Then SendLine "CUPOPEN|" & gCupId(k)
    Case "cupdel"
        If MsgBox("Delete " & gCupName(k) & "? Its bracket and results are removed from every PC.", vbQuestion + vbOKCancel + vbDefaultButton2, "Delete Tournament") = vbOK Then SendLine "CUPDEL|" & gCupId(k)
    End Select
End Sub

' Report Result (a player, for their own match), or Enter/Confirm Result (the host, for any match)
Private Sub CupResultDialog(ByVal m As Long)
    Dim k As Long, r As Long, i As Long, a As String, b As String
    k = CupK()
    If k < 0 Or m < 0 Then Exit Sub
    r = m \ 100
    i = m Mod 100
    a = CupM(k, r, i, "a")
    b = CupM(k, r, i, "b")
    If Len(a) = 0 Or a = "-" Or Len(b) = 0 Or b = "-" Then Exit Sub
    If gCupState(k) <> "running" And gCupState(k) <> "done" Then Exit Sub
    If Not mIsAdmin Then
        If StrComp(a, mMyNick, vbTextCompare) <> 0 And StrComp(b, mMyNick, vbTextCompare) <> 0 Then Exit Sub
        If Len(CupM(k, r, i, "win")) > 0 Then
            Note "The result of this match has been confirmed. To change it, ask the host."
            Exit Sub
        End If
    End If
    frmCupResult.ShowMatch gCupId(k), r, i
End Sub

' Copy to Chat: the match picked, or how the tournament stands, typed into the chat box ready to send
Private Sub CupToChat(ByVal k As Long)
    Dim s As String, parts() As String
    If mCupMatch >= 0 Then
        parts = Split(CupMatchText(k, mCupMatch \ 100, mCupMatch Mod 100, "", False) & vbLf, vbLf)
        s = gCupName(k) & ", " & parts(0) & " - " & parts(1)
    ElseIf Len(gCupChamp(k)) > 0 Then
        s = gCupChamp(k) & " won " & gCupName(k) & "."
    Else
        s = gCupName(k) & ": " & CupStatusText(k, "")
    End If
    GoPage PG_CHAT
    txtSay.Text = Left$(s, txtSay.MaxLength)
    txtSay.SelStart = Len(txtSay.Text)
    On Error Resume Next
    txtSay.SetFocus
End Sub

' ---- news from the server

' A tournament changed (or went): keep the page, the toolbar and the sidebar up to date
Private Sub CupChanged()
    Dim k As Long
    RefreshOpenForm "frmProjector"
    If CupK() < 0 Then
        mCupMatch = -1
        If gCupCount > 0 Then
            mCupSel = gCupId(0)
        Else
            mCupSel = ""
        End If
    End If
    k = CupK()
    If k >= 0 And mCupMatch >= 0 Then
        If mCupMatch \ 100 >= gCupRounds(k) Then
            mCupMatch = -1
        ElseIf mCupMatch Mod 100 >= CupMatches(k, mCupMatch \ 100) Then
            mCupMatch = -1
        End If
    End If
    UpdateCupScroll
    DrawCupPage
    If mPage = PG_CUPS Then
        DrawHeader
        BuildTools
        DrawTools
    End If
    CupNavBadge
    If IsFormLoaded("frmCupPlayers") Then frmCupPlayers.RefreshCup
End Sub

' A message from the server about a tournament command that didn't work
Private Sub CupError(ByVal msg As String)
    Beep
    If IsFormLoaded("frmCupPlayers") Then
        frmCupPlayers.ShowError msg
    ElseIf mPage = PG_CUPS Then
        mCupError = msg
        DrawCupPage
    Else
        AddSys IC_WARN, msg, C_RED
    End If
End Sub

' Sign-ups for a new tournament: a pop-up for everyone
Private Sub CupSignupNews(ByVal id As String)
    Dim k As Long, s As String
    k = CupIndex(id)
    If k < 0 Or Not gNotifyCups Then Exit Sub
    mCupToast = id
    mCupToastMatch = -1
    s = gCupName(k)
    If Len(gCupGame(k)) > 0 Then s = s & " (" & gCupGame(k) & ")"
    Ding
    ShowToast "cup", IC_TROPHY, "Tournament Sign-ups Open", s & ". Click here to sign up.", 10
End Sub

' Your next match has both players: time to play it
Private Sub CupMatchNews(ByVal id As String, ByVal r As Long, ByVal i As Long, ByVal opp As String)
    Dim k As Long
    k = CupIndex(id)
    If k < 0 Then Exit Sub
    FlashTitle Me.hWnd
    If Not gNotifyCups Then Exit Sub
    mCupToast = id
    mCupToastMatch = r * 100 + i
    PlayChime "ding"
    ShowToast "cupmatch", IC_FLAG, "Your Match Is Ready", gCupName(k) & ", " & CupRoundWord(k, r) & ": you against " & opp & ". Click here for the bracket.", 15
End Sub

' The host: a player reported a result
Private Sub CupReportNews(ByVal id As String, ByVal r As Long, ByVal i As Long, ByVal who As String)
    Dim k As Long
    k = CupIndex(id)
    If k < 0 Or Not mIsAdmin Or Not gNotifyCups Then Exit Sub
    mCupToast = id
    mCupToastMatch = r * 100 + i
    Ding
    ShowToast "cupreport", IC_WARN, "Result Reported", who & " reported a result in " & gCupName(k) & " (" & CupRoundWord(k, r) & "). Click here to confirm it.", 12
End Sub

Private Sub CupWinNews(ByVal id As String, ByVal nick As String)
    Dim k As Long
    k = CupIndex(id)
    If k < 0 Or Not gNotifyCups Then Exit Sub
    mCupToast = id
    mCupToastMatch = -1
    If StrComp(nick, mMyNick, vbTextCompare) = 0 Then
        PlayChime "ding"
        ShowToast "cup", IC_TROPHY, "You Won " & gCupName(k), "Congratulations. You are the winner of " & gCupName(k) & ".", 15
    Else
        Ding
        ShowToast "cup", IC_TROPHY, "Tournament Winner", nick & " won " & gCupName(k) & ".", 10
    End If
End Sub

' The red bubble on Tournaments in the sidebar: matches waiting for you (and, for the host, results to confirm)
Private Sub CupNavBadge()
    Dim k As Long, n As Long, s As String
    For k = 0 To gCupCount - 1
        If CupMyTurn(k, mMyNick) Then n = n + 1
        If mIsAdmin Then n = n + CupPending(k)
    Next
    If n > 9 Then
        s = "9+"
    ElseIf n > 0 Then
        s = CStr(n)
    End If
    If s <> mCupBadge Then
        mCupBadge = s
        DrawSide
    End If
End Sub

' ---- for the tournament windows

Public Function MyNick() As String
    MyNick = mMyNick
End Function

' Everyone signed in, us included, separated by |
Public Function PeopleHere() As String
    Dim i As Long, s As String
    For i = 0 To mWhoCount - 1
        s = s & mWhoNick(i) & "|"
    Next
    PeopleHere = s
End Function

Public Sub CupNote(ByVal s As String)
    Note s
End Sub

' ================================================================ Make Cartridge Wizard

' Tools > Make Cartridge, and the button on the Game Library toolbar. The wizard (frmCartWizard) stays
' open beside the main window while 7-Zip packs the game; a second click brings it to the front.
Private Sub ShowCartWizard()
    On Error Resume Next
    If IsFormLoaded("frmCartWizard") Then
        If frmCartWizard.WindowState = vbMinimized Then frmCartWizard.WindowState = vbNormal
        frmCartWizard.ZOrder 0
        frmCartWizard.SetFocus
    Else
        frmCartWizard.Show vbModeless, Me
    End If
End Sub

' The wizard is making a cartridge
Private Function CartBuilding() As Boolean
    If IsFormLoaded("frmCartWizard") Then CartBuilding = frmCartWizard.IsBuilding()
End Function

' ================================================================ pop-up notifications

Public Sub ShowToast(ByVal kind As String, ByVal ic As Integer, ByVal title As String, ByVal msg As String, ByVal secs As Long)
    Dim t As frmToast, a As frmAlert, i As Integer, n As Long
    On Error Resume Next
    ' Busy: only announcements from the host, Remind Me (and the Test button) get through
    If mBusy And kind <> "pa" And kind <> "test" And kind <> "remind" And kind <> "cupmatch" Then Exit Sub
    If gClassicPopups Then
        ' the classic box: count the ones already open so the new one sits a little lower
        For i = 0 To Forms.Count - 1
            If Forms(i).Name = "frmAlert" Then n = n + 1
        Next
        If n >= 3 Then Exit Sub
        Set a = New frmAlert
        a.ShowAlert kind, ic, title, msg, n
    Else
        ' four at once is plenty: any more would pile up on top of each other
        If gToastSlot(0) And gToastSlot(1) And gToastSlot(2) And gToastSlot(3) Then Exit Sub
        Set t = New frmToast
        t.Popup kind, ic, title, msg, secs
    End If
End Sub

' Someone clicked a pop-up: bring the LAN Center up at the right page
Public Sub ToastClicked(ByVal kind As String)
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    Me.ZOrder 0
    Select Case kind
    Case "vote", "soon", "result", "next", "remind"
        GoPage PG_WHATSON
    Case "food"
        GoPage PG_WHATSON
        frmFood.Show vbModal, Me
    Case "pa", "name"
        GoPage PG_CHAT
    Case "lib"
        GoPage PG_LIBRARY
    Case "cup", "cupmatch", "cupreport"
        GoPage PG_CUPS
        If CupIndex(mCupToast) >= 0 Then
            SelectCup mCupToast
            mCupMatch = mCupToastMatch
            CupScrollTo mCupMatch
            DrawCupPage
            RefreshCupTools
        End If
    Case "libnotes"
        If LibShown(mLibToastId) Then mLibSel = mLibToastId
        GoPage PG_LIBRARY
        LibShowNotes mLibToastId
        DrawLibCard
    End Select
    Me.SetFocus
End Sub

' Called by the Settings window after OK
Public Sub SettingsChanged()
    mnuToolsSounds.Checked = gSoundOn
    BuildTools
    DrawTools
    DrawLibPage           ' the games folder may have changed
End Sub

' Called by the Mute Players window after OK
Public Sub IgnoreChanged()
    Note "Muted players updated."
    DrawWho
End Sub

' Everyone signed in apart from us, for the Mute Players window
Public Function OtherPeople() As String
    Dim i As Long, s As String
    For i = 0 To mWhoCount - 1
        If StrComp(mWhoNick(i), mMyNick, vbTextCompare) <> 0 Then s = s & mWhoNick(i) & "|"
    Next
    OtherPeople = s
End Function

' ---- Help > Diagnostics. Lines of "[Heading]" or "Label" & vbTab & "Value".

Private Function DiagRow(ByVal lbl As String, ByVal v As String) As String
    If Len(v) = 0 Then v = "-"
    DiagRow = lbl & vbTab & v & vbCrLf
End Function

Private Function SocketWord() As String
    Select Case sckTCP.State
    Case sckClosed: SocketWord = "Not connected"
    Case sckResolvingHost, sckHostResolved, sckConnecting: SocketWord = "Connecting"
    Case sckConnected: SocketWord = "Connected"
    Case sckClosing: SocketWord = "Disconnecting"
    Case sckError: SocketWord = "Error"
    Case Else: SocketWord = "State " & sckTCP.State
    End Select
End Function

' The parts that change: the servers, and the Game Library
Public Function DiagLive() As String
    Dim s As String, t As String, free As Double, missing As String
    s = "[Palms Server]" & vbCrLf
    s = s & DiagRow("Name", mServerName)
    If Len(mServerHost) > 0 Then t = mServerHost & ":" & mServerPort Else t = "Not set"
    s = s & DiagRow("Address", t)
    s = s & DiagRow("Version", mServerVer)
    s = s & DiagRow("Connection", SocketWord())
    s = s & DiagRow("Last status", mConnText)
    If mPingMs >= 0 And sckTCP.State = sckConnected Then t = mPingMs & " ms" Else t = "-"
    s = s & DiagRow("Round trip", t)
    If mSignedIn Then
        t = mMyNick
        If Len(mMySeat) > 0 Then t = t & ", seat " & mMySeat
        t = t & ", " & MyStatusWord()
        If mIsAdmin Then t = t & ", host mode"
    Else
        t = "No"
    End If
    s = s & DiagRow("Signed in", t)
    If mFindBroken Then t = "Not available" Else t = "Listening on UDP port " & BEACON_PORT
    s = s & DiagRow("Server discovery", t)
    s = s & "[Cartridge Server]" & vbCrLf
    If Len(gPakBase) > 0 Then t = gPakBase Else t = "Not received from the Palms Server"
    s = s & DiagRow("Address", t)
    If Not gPakHeard Then
        t = "Waiting for the game list"
    ElseIf Len(gPakNote) > 0 Then
        t = gPakNote
    Else
        t = "OK"
    End If
    s = s & DiagRow("Status", t)
    s = s & DiagRow("Cartridges", CStr(gPakCount))
    s = s & "[Game Library]" & vbCrLf
    t = GamesRoot()
    free = FreeBytes(GamesRoot())
    If free >= 0 Then t = t & " (" & SizeText(free) & " free)"
    s = s & DiagRow("Games folder", t)
    missing = ToolsMissing()
    If Len(missing) > 0 Then t = "Missing: " & missing Else t = "palmsget.exe and 7za.exe found"
    s = s & DiagRow("Helper programs", t)
    s = s & DiagRow("Installed", CStr(gHereCount))
    If Len(gJobId) > 0 Then t = gJobText Else t = "None"
    s = s & DiagRow("Current task", t)
    If Len(gFailText) > 0 Then t = gFailText Else t = "None"
    s = s & DiagRow("Last error", t)
    DiagLive = s
End Function

' The parts that don't change: this PC
Public Function DiagPC() As String
    Dim s As String, nic As String, mbps As Long, t As String, ip As String, sip As String
    s = "[This PC]" & vbCrLf
    s = s & DiagRow("Computer name", ThisPCName())
    On Error Resume Next
    ip = sckTCP.LocalIP
    On Error GoTo 0
    s = s & DiagRow("IP address", ip)
    s = s & DiagRow("PC ID", mPcId)
    s = s & DiagRow("Palms LAN Center", AppVersion())
    s = s & DiagRow("Windows", WindowsName())
    s = s & DiagRow("Processor", CpuName())
    s = s & DiagRow("Memory", MemoryText())
    If sckTCP.State = sckConnected Then sip = sckTCP.RemoteHostIP Else sip = mServerHost
    NetworkCard sip, nic, mbps
    t = nic
    If mbps > 0 Then t = t & ", " & mbps & " Mbps"
    s = s & DiagRow("Network card", t)
    s = s & DiagRow("Display", ScreenText())
    DiagPC = s
End Function



' ================================================================ the mouse wheel (in PalmsLAN.exe only)

Public Sub OnWheel(ByVal delta As Long)
    Dim hw As Long, steps As Long
    On Error Resume Next
    If delta = 0 Then Exit Sub
    steps = delta \ 120
    If steps = 0 Then steps = Sgn(delta)
    steps = -3 * steps
    hw = WindowUnderMouse()
    If hw = picWho.hWnd Or hw = vsbWho.hWnd Then
        ScrollBy vsbWho, steps
    ElseIf hw = picSched.hWnd Or hw = vsbSched.hWnd Then
        ScrollBy vsbSched, steps
    ElseIf hw = picPeople.hWnd Or hw = vsbPeople.hWnd Then
        ScrollBy vsbPeople, steps
    ElseIf hw = picMap.hWnd Or hw = vsbMap.hWnd Then
        ScrollBy vsbMap, steps * 20
    ElseIf hw = picLibList.hWnd Or hw = vsbLib.hWnd Then
        ScrollBy vsbLib, steps
    ElseIf hw = hsbCup.hWnd Then
        SetCupScroll CupSX() + steps * 20, CupSY()
    ElseIf hw = picCup.hWnd Or hw = vsbCup.hWnd Then
        If vsbCup.Visible Then
            SetCupScroll CupSX(), CupSY() + steps * 20
        Else
            SetCupScroll CupSX() + steps * 20, CupSY()
        End If
    ElseIf hw = picSide.hWnd Then
        NavScroll Sgn(steps)
    ElseIf mPage = PG_CHAT Then
        ScrollLog steps
    End If
End Sub

Private Sub ScrollBy(sb As VScrollBar, ByVal delta As Long)
    Dim v As Long
    v = sb.Value + delta
    If v < 0 Then v = 0
    If v > sb.Max Then v = sb.Max
    sb.Value = v
End Sub

' ================================================================ updating itself

' ================================================================ finding and connecting

' Listen for the server's "I'm here" shout on UDP port 4099
Private Sub StartListening()
    On Error GoTo Failed
    sckFind.Protocol = sckUDPProtocol
    sckFind.Bind BEACON_PORT
    Exit Sub
Failed:
    lblFind.Caption = "Unable to listen for server broadcasts (" & Err.Description & "). Enter the server address below."
End Sub

Private Sub sckFind_DataArrival(ByVal bytesTotal As Long)
    Dim s As String, f() As String
    On Error Resume Next
    sckFind.GetData s, vbString
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    f = Split(s, "|")
    Select Case f(0)
    Case "PPING"
        ' another LAN Center measuring its ping to us: answer straight away
        If UBound(f) >= 1 Then SendUdp sckFind.RemoteHostIP, "PPONG|" & f(1)
        Exit Sub
    Case "PPONG"
        ' the answer to one of ours
        If UBound(f) >= 1 Then PeerAnswered sckFind.RemoteHostIP, Val(f(1))
        Exit Sub
    End Select
    If UBound(f) < 2 Then Exit Sub
    If f(0) <> "PALMS" Then Exit Sub
    If IsBusy() Then Exit Sub
    mServerHost = sckFind.RemoteHostIP
    mServerPort = Val(f(2))
    If mServerPort <= 0 Then mServerPort = DEFAULT_PORT
    txtServer.Text = mServerHost
    If Not mSignedIn Then lblFind.Caption = "Found " & f(1) & " at " & mServerHost & "."
    ConnectTo mServerHost
End Sub

Private Sub sckFind_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    mFindBroken = True
End Sub

Private Function IsBusy() As Boolean
    Select Case sckTCP.State
    Case sckResolvingHost, sckHostResolved, sckConnecting, sckConnected
        IsBusy = True
    End Select
End Function

Private Sub ConnectTo(ByVal host As String)
    On Error GoTo Failed
    tmrReconnect.Enabled = False
    If sckTCP.State <> sckClosed Then sckTCP.Close
    mBuffer = ""
    mServerHost = host
    SetConn IC_REFRESH, "Connecting to " & host & "..."
    sckTCP.RemoteHost = host
    sckTCP.RemotePort = mServerPort
    sckTCP.Connect
    Exit Sub
Failed:
    SetConn IC_WARN, "Unable to connect to " & host & ": " & Err.Description
    tmrReconnect.Enabled = True
End Sub

Private Sub sckTCP_Connect()
    SetConn IC_OK, "Connected to " & mServerHost & "."
    SendLine "HELLO|" & mPcId & "|" & AppVersion() & "|" & ThisPCName()
End Sub

Private Sub sckTCP_Close()
    sckTCP.Close
    If mClosing Then Exit Sub
    LostServer
End Sub

Private Sub sckTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    sckTCP.Close
    If mClosing Then Exit Sub
    SetConn IC_WARN, "Unable to connect to " & mServerHost & " (" & Description & "). Retrying..."
    If Not mSignedIn Then lblFind.Caption = "Unable to connect to " & mServerHost & ". Check the server address."
    LostServer
End Sub

Private Sub LostServer()
    mIsAdmin = False
    mPingMs = -1
    If mSigningOut Then
        mSigningOut = False
        mAutoRelogin = False
        mMySeat = ""
        SettingSet "seat", ""
        txtSeat.Text = ""
        ShowSignIn
        lblFind.Caption = "Signed out."
        SetConn IC_INFO, "Signed out."
    ElseIf mSignedIn Then
        AddSys IC_WARN, "Connection to the server lost. Reconnecting...", C_RED
        SetConn IC_WARN, "Connection lost. Reconnecting..."
        mAutoRelogin = True
        mSignedIn = False
        If mPage = PG_PA Or mPage = PG_SIGNIN Then
            ShowPage PG_CHAT
        Else
            ShowPage mPage
        End If
    Else
        cmdSignIn.Enabled = True
    End If
    UpdateMenus
    DrawStatus
    tmrReconnect.Enabled = True
End Sub

Private Sub tmrReconnect_Timer()
    If sckTCP.State = sckConnected Then
        tmrReconnect.Enabled = False
        Exit Sub
    End If
    If IsBusy() Then Exit Sub
    If Len(mServerHost) > 0 Then ConnectTo mServerHost
End Sub

Private Sub tmrPing_Timer()
    SendPing
End Sub

Private Sub SendPing()
    If sckTCP.State = sckConnected Then
        mPingSent = GetTickCount()
        SendLine "PING"
    End If
End Sub

Public Sub SendLine(ByVal text As String)
    On Error Resume Next
    If sckTCP.State = sckConnected Then sckTCP.SendData text & vbCrLf
End Sub

' ================================================================ messages from the server

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
    Dim s As String, p As Long, lineText As String
    On Error Resume Next
    sckTCP.GetData s, vbString
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    mBuffer = mBuffer & s
    p = InStr(mBuffer, vbLf)
    Do While p > 0
        lineText = Left$(mBuffer, p - 1)
        mBuffer = Mid$(mBuffer, p + 1)
        If Right$(lineText, 1) = vbCr Then lineText = Left$(lineText, Len(lineText) - 1)
        If Len(lineText) > 0 Then SafeHandleLine lineText
        p = InStr(mBuffer, vbLf)
    Loop
    If Len(mBuffer) > 32768 Then mBuffer = ""
End Sub

' Never let one odd message crash the program
Private Sub SafeHandleLine(ByVal lineText As String)
    On Error GoTo Oops
    PerfStart PF_MSG
    HandleLine lineText
    PerfStop PF_MSG
    Exit Sub
Oops:
    PerfStop PF_MSG
    Note "Invalid message from the server (" & Err.Description & ")."
End Sub

Private Sub HandleLine(ByVal lineText As String)
    Dim f() As String, cmd As String, msg As String, who As String, s As String
    f = Split(lineText, "|")
    cmd = UCase$(f(0))
    Select Case cmd
    Case "SPING"
        ' the server measuring our ping: answer straight away
        SendLine "SPONG|" & Fld(f, 1)

    Case "WELCOME"
        mServerName = Fld(f, 1)
        mServerVer = Fld(f, 2)
        If Not mSignedIn Then lblFind.Caption = "Connected to " & mServerName & " at " & mServerHost & "."
        SetConn IC_OK, "Connected to " & mServerName & " (" & mServerHost & ")."
        SettingSet "server", mServerHost
        SendPing
        If mAutoRelogin And Len(mMyNick) > 0 Then
            SendLine "LOGIN|" & mMyNick & "|" & SeatOrDash(mMySeat)
        ElseIf mPendingLogin Then
            SendLine "LOGIN|" & mWantNick & "|" & SeatOrDash(mWantSeat)
        End If

    Case "KNOWN"
        If Len(txtNick.Text) = 0 Then txtNick.Text = Fld(f, 1)
        If Len(txtSeat.Text) = 0 And Fld(f, 2) <> "-" Then txtSeat.Text = Fld(f, 2)
        If Not mSignedIn And Not mPendingLogin And Not mAutoRelogin Then
            lblFind.Caption = "Welcome back, " & Fld(f, 1) & ". Click Sign In to continue."
        End If

    Case "NEWPC"
        If Not mSignedIn And Not mPendingLogin Then
            lblFind.Caption = "Connected. Enter a nickname and click Sign In."
        End If

    Case "PONG"
        If mPingSent <> 0 Then
            mPingMs = GetTickCount() - mPingSent
            If mPingMs < 0 Then mPingMs = 0
        End If
        DrawStatus

    Case "LOGGEDIN"
        mPendingLogin = False
        mAutoRelogin = False
        mSignedIn = True
        mMyNick = Fld(f, 1)
        mMySeat = Fld(f, 2)
        If mMySeat = "-" Then mMySeat = ""
        mAway = False
        mBusy = False
        mIsAdmin = (Fld(f, 3) = "1")
        mSuggFresh = True
        mSuggCount = 0
        mFoodHeard = False
        mFoodState = "closed"
        mFoodCount = 0
        CupClear
        mCupSel = ""
        mCupMatch = -1
        mCupHeard = False
        mCupBadge = ""
        SettingSet "nick", mMyNick
        SettingSet "seat", mMySeat
        ClearLog
        ShowPage PG_CHAT
        ' said after the recent chat has come in (see FinishWho)
        If Len(mMySeat) > 0 Then
            mWelcome = "Signed in as " & mMyNick & " at seat " & mMySeat & "."
        Else
            mWelcome = "Signed in as " & mMyNick & "."
            If Len(mWelcomeNote) = 0 Then mWelcomeNote = "To select a seat, click your name on the toolbar and click Select Seat."
        End If
        SetConn IC_OK, "Connected to the Palms Server."
        RequestWho
        ' share this PC's setup for the Who's Here page, and let the server measure our ping
        If ServerIsNew() Then
            If ServerHasCups() Then
                SendLine "CAPS|ping|paks|cups"
            ElseIf ServerHasPaks() Then
                SendLine "CAPS|ping|paks"
            Else
                SendLine "CAPS|ping"
            End If
            SendLine "SPECS|" & SpecsLine(sckTCP.RemoteHostIP)
            ' the Game Library's list of games
            If ServerHasPaks() Then SendLine "PAKS"
            ' the tournaments
            If ServerHasCups() Then SendLine "CUPS"
        End If

    Case "HIST"
        If Not IsIgnored(Fld(f, 2)) Then AddLog ShortTime(Fld(f, 1)), -1, Fld(f, 2), NickColor(Fld(f, 2)), RestOf(f, 3), vbBlack, -1

    Case "CHAT"
        who = Fld(f, 2)
        msg = RestOf(f, 3)
        If Not IsIgnored(who) Then AddLog ShortTime(Fld(f, 1)), -1, who, NickColor(who), msg, vbBlack, -1
        If StrComp(who, mMyNick, vbTextCompare) <> 0 And Not IsIgnored(who) Then
            If mPage <> PG_CHAT Then
                mUnread = mUnread + 1
                DrawSide
            End If
            If gNotifyName And InStr(1, msg, mMyNick, vbTextCompare) > 0 Then
                Ding
                FlashTitle Me.hWnd
                ' no need to pop up if you're looking at the chat right now
                If GetForegroundWindow() <> Me.hWnd Or mPage <> PG_CHAT Then
                    ' at most one of these every 5 seconds, however chatty people are
                    If mLastNameToast = 0 Or GetTickCount() - mLastNameToast > 5000 Then
                        mLastNameToast = GetTickCount()
                        ShowToast "name", IC_CHAT, who & " says:", msg, 8
                    End If
                End If
            End If
        End If

    Case "JOIN"
        ' back after updating: nothing to say
        If gShowJoins And Fld(f, 3) <> "update" Then AddSys IC_DOOR, Fld(f, 1) & " has signed in" & SeatText(Fld(f, 2)) & ".", C_GREEN
        RequestWho

    Case "LEFT"
        If gShowJoins Then
            Select Case Fld(f, 2)
            Case "home"
                AddSys IC_DOOR, Fld(f, 1) & " has signed out.", C_GREY
            Case "quit"
                AddSys IC_DOOR, Fld(f, 1) & " has exited Palms LAN Center.", C_GREY
            Case "update"
                ' updating Palms LAN Center: back in a few seconds
            Case Else
                AddSys IC_WARN, Fld(f, 1) & " has been disconnected.", C_RED
            End Select
        End If
        RequestWho

    Case "STATUS"
        If gShowJoins Then
            Select Case Fld(f, 2)
            Case "away"
                AddSys IC_MOON, Fld(f, 1) & " is away.", C_GREY
            Case "busy"
                AddSys IC_BAD, Fld(f, 1) & " is busy.", C_GREY
            Case Else
                AddSys IC_DOOR, Fld(f, 1) & " is online.", C_GREEN
            End Select
        End If
        If StrComp(Fld(f, 1), mMyNick, vbTextCompare) = 0 Then
            mAway = (Fld(f, 2) = "away")
            mBusy = (Fld(f, 2) = "busy")
            UpdateMenus
            BuildTools
            DrawTools
        End If
        RequestWho

    Case "WHO"
        If mWhoBuild Is Nothing Then Set mWhoBuild = New Collection
        mWhoBuild.Add lineText

    Case "GONE"
        If mGoneBuild Is Nothing Then Set mGoneBuild = New Collection
        mGoneBuild.Add lineText

    Case "MAP"
        ' the seat map in use (sent when you sign in, and when the host changes it)
        MapFromFields f, 1, mMap
        If mPage = PG_WHO And mWhoView = 1 Then DrawMapView

    Case "MAPL", "MAPLEND", "MAPDEF", "MAPSAVED"
        If IsFormLoaded("frmMapEdit") Then frmMapEdit.OnServer lineText

    Case "WHOEND"
        FinishWho

    Case "PAKS"
        ' the Game Library: the games on the pak server, sent when you sign in and when they change
        PaksBegin f

    Case "PAK"
        PaksAdd f

    Case "GAME"
        ' one game in a pack, sent after its PAK line
        PaksGame f

    Case "PAKEND"
        PaksEnd
        LibRefresh
        LibFetchIcons

    Case "WHOIS"
        If StrComp(Fld(f, 1), mPplSel, vbTextCompare) = 0 Then
            mCard = f
            mCardNick = Fld(f, 1)
            If mPage = PG_WHO Then
                DrawCard
                UpdateWhoButtons
            End If
        End If

    Case "SEATCHG"
        s = Fld(f, 2)
        If s = "-" Then s = ""
        If StrComp(Fld(f, 1), mMyNick, vbTextCompare) = 0 Then
            mMySeat = s
            SettingSet "seat", mMySeat
            txtSeat.Text = mMySeat
            If IsFormLoaded("frmSeat") Then Unload frmSeat
            If Len(s) > 0 Then
                AddSys IC_OK, "Your seat is now " & s & ".", C_GREEN
            Else
                AddSys IC_INFO, "You no longer have a seat.", C_GREY
            End If
            DrawTools
        ElseIf Len(s) > 0 Then
            AddSys IC_USER, Fld(f, 1) & " moved to seat " & s & ".", C_GREY
        End If
        RequestWho

    Case "TPL", "TPLEND", "TPLTEXT", "TPLSAVED", "TPLAPPLIED"
        If IsFormLoaded("frmTemplates") Then frmTemplates.OnServer lineText

    Case "SCHED"
        If mSchedBuild Is Nothing Then Set mSchedBuild = New Collection
        mSchedBuild.Add lineText

    Case "SCHEDEND"
        FinishSched

    Case "SCHEDINFO", "SCHEDGAME"
        If mSchedInfo Is Nothing Then Set mSchedInfo = New Collection
        mSchedInfo.Add lineText

    Case "PROJINFO"
        mProjSecs = Val(Fld(f, 1))
        mProjSlides = Fld(f, 2)
        mProjPort = Fld(f, 3)
        RefreshOpenForm "frmProjector"

    Case "SCHEDADDED"
        ' the event the host just added: now its details can go with it
        If Len(mPendingNote) > 0 Then
            SendLine "SCHEDNOTE|" & Fld(f, 1) & "|" & mPendingNote
            mPendingNote = ""
        End If
        If Len(mPendingGame) > 0 Then
            SendLine "SCHEDGAME|" & Fld(f, 1) & "|" & mPendingGame
            mPendingGame = ""
        End If

    Case "SUGG"
        If mSuggFresh Then
            mSuggCount = 0
            mSuggFresh = False
        End If
        If mSuggCount <= 15 Then
            mSuggN(mSuggCount) = Val(Fld(f, 1))
            mSuggBy(mSuggCount) = Fld(f, 2)
            mSuggName(mSuggCount) = RestOf(f, 3)
            mSuggCount = mSuggCount + 1
        End If

    Case "SUGGEND"
        If mSuggFresh Then mSuggCount = 0
        mSuggFresh = True
        RefreshOpenForm "frmSuggest"
        RefreshOpenForm "frmVote"

    Case "FOODINFO"
        FoodNews Fld(f, 1), RestOf(f, 3)
        mFoodState = Fld(f, 1)
        mFoodTitle = RestOf(f, 3)
        mFoodCount = 0

    Case "FOODITEM"
        If mFoodCount <= 99 Then
            mFoodNick(mFoodCount) = Fld(f, 1)
            mFoodTime(mFoodCount) = Fld(f, 2)
            mFoodText(mFoodCount) = RestOf(f, 3)
            mFoodCount = mFoodCount + 1
        End If

    Case "FOODEND"
        RefreshOpenForm "frmFood"

    Case "VOTEINFO"
        mPrevVoteState = mVoteState
        mVoteState = Fld(f, 1)
        mVoteDeadline = GetTickCount() + Val(Fld(f, 2)) * 1000
        mVoteMine = Val(Fld(f, 3))
        mVoteQ = RestOf(f, 4)
        mVoteCount = 0

    Case "VOTEOPT"
        If mVoteCount <= 5 Then
            mVoteVotes(mVoteCount) = Val(Fld(f, 2))
            mVoteName(mVoteCount) = RestOf(f, 3)
            mVoteCount = mVoteCount + 1
        End If

    Case "VOTEEND"
        mVoteTotal = Val(Fld(f, 1))
        mVoteWinner = Val(Fld(f, 2))
        VoteNews
        DrawVote
        DrawNext
        UpdateMenus
        RefreshWhatsTools
        RefreshOpenForm "frmProjector"

    Case "CUPCLEAR"
        ' all the tournaments follow: after them, a new one gets a pop-up
        CupLine lineText
        mCupWait = Val(Fld(f, 1))
        mCupHeard = (mCupWait = 0)
        CupChanged

    Case "CUP"
        s = ""
        If mCupHeard And CupIndex(Fld(f, 1)) < 0 And Fld(f, 2) = "signup" Then s = Fld(f, 1)
        CupLine lineText
        If Len(s) > 0 Then mCupNew = s

    Case "CUPNOTE", "CUPP", "CUPM"
        CupLine lineText

    Case "CUPEND"
        CupLine lineText
        If mCupWait > 0 Then
            mCupWait = mCupWait - 1
            If mCupWait = 0 Then mCupHeard = True
        End If
        CupChanged
        If Len(mCupNew) > 0 Then
            CupSignupNews mCupNew
            mCupNew = ""
        End If

    Case "CUPDEL"
        CupLine lineText
        CupChanged

    Case "CUPSAVED"
        ' the host's new or changed tournament: show it
        GoPage PG_CUPS
        SelectCup Fld(f, 1)

    Case "CUPMATCH"
        CupMatchNews Fld(f, 1), Val(Fld(f, 2)), Val(Fld(f, 3)), Fld(f, 4)

    Case "CUPREPORTED"
        CupReportNews Fld(f, 1), Val(Fld(f, 2)), Val(Fld(f, 3)), Fld(f, 4)

    Case "CUPWIN"
        CupWinNews Fld(f, 1), Fld(f, 2)

    Case "NOTE"
        AddSys EventIcon(Fld(f, 1)), RestOf(f, 2), C_BLUE
        If mPage <> PG_CHAT Then
            mUnread = mUnread + 1
            DrawSide
        End If
        FlashTitle Me.hWnd

    Case "ADMINOK"
        mIsAdmin = True
        UpdateMenus
        DrawSide
        BuildTools
        DrawTools
        DrawWhats
        AddSys IC_ADMIN, "Host mode enabled. The Host menu and the Announce page are now available.", C_PURPLE
        DrawCupPage
        CupNavBadge
        RequestWho

    Case "PA"
        AddLog ShortTime(Fld(f, 3)), IC_MEGAPHONE, "", 0, "Announcement from " & Fld(f, 2) & ": " & RestOf(f, 4), vbBlack, C_TIP
        If mPage <> PG_CHAT Then
            mUnread = mUnread + 1
            DrawSide
        End If
        If gNotifyPA Then
            PlayChime Fld(f, 1)
            ShowToast "pa", IC_MEGAPHONE, "Announcement from " & Fld(f, 2), RestOf(f, 4), 15
        End If
        FlashTitle Me.hWnd

    Case "ERR"
        If Not mSignedIn Then
            mPendingLogin = False
            If mAutoRelogin Then
                mAutoRelogin = False
                ShowSignIn
            End If
            lblFind.Caption = RestOf(f, 2)
            cmdSignIn.Enabled = True
            Beep
        ElseIf IsFormLoaded("frmTemplates") And (Fld(f, 1) = "NO_TIMES" Or Fld(f, 1) = "NOT_FOUND") Then
            frmTemplates.OnServer lineText
        ElseIf IsFormLoaded("frmMapEdit") And (Left$(Fld(f, 1), 4) = "MAP_" Or Fld(f, 1) = "NOT_FOUND" Or Fld(f, 1) = "NOT_ADMIN") Then
            frmMapEdit.OnServer lineText
        ElseIf IsFormLoaded("frmSuggest") And (Fld(f, 1) = "TOO_MANY" Or Fld(f, 1) = "FULL" Or Fld(f, 1) = "EMPTY") Then
            frmSuggest.ShowError RestOf(f, 2)
        ElseIf IsFormLoaded("frmFood") And Fld(f, 1) = "FOOD_CLOSED" Then
            frmFood.ShowError RestOf(f, 2)
        ElseIf Left$(Fld(f, 1), 4) = "CUP_" Then
            CupError RestOf(f, 2)
        ElseIf Fld(f, 1) = "SEAT_TAKEN" Or Fld(f, 1) = "BAD_SEAT" Then
            If IsFormLoaded("frmSeat") Then
                frmSeat.ShowError RestOf(f, 2)
            ElseIf Len(mWelcome) > 0 Then
                ' signing in: the server let us in without the seat. Say so after the welcome.
                mWelcomeNote = RestOf(f, 2) & " To select a seat, click your name on the toolbar."
            Else
                AddSys IC_WARN, RestOf(f, 2), C_RED
            End If
        Else
            AddSys IC_WARN, RestOf(f, 2), C_RED
        End If

    Case "BYE"
        If mSignedIn Then AddSys IC_INFO, RestOf(f, 1), C_GREY

    Case "INFO"
        If mSignedIn Then AddSys IC_INFO, RestOf(f, 1), C_GREY
    End Select
End Sub

' ================================================================ buttons and menus

Private Sub cmdSignIn_Click()
    Dim nick As String, seat As String, host As String
    nick = Trim$(txtNick.Text)
    seat = UCase$(Trim$(txtSeat.Text))
    host = Trim$(txtServer.Text)
    If Not ValidNick(nick) Then
        MsgBox "Nicknames must be 2 to 12 characters long and can contain letters, numbers, hyphens (-) and underscores (_).", vbExclamation, "Sign In"
        txtNick.SetFocus
        Exit Sub
    End If
    ' a seat that doesn't look right never stops you signing in: you start without one
    mWelcomeNote = ""
    If Len(seat) > 0 And Not ValidSeat(seat) Then
        mWelcomeNote = seat & " is not a valid seat number (for example, A01), so no seat was selected. To select a seat, click your name on the toolbar."
        seat = ""
    End If
    If Len(host) = 0 Then
        MsgBox "The Palms Server was not found. Enter the server address (for example, 10.0.200.71).", vbExclamation, "Sign In"
        txtServer.SetFocus
        Exit Sub
    End If
    mWantNick = nick
    mWantSeat = seat
    cmdSignIn.Enabled = False
    lblFind.Caption = "Signing in..."
    If sckTCP.State = sckConnected And host = mServerHost Then
        SendLine "LOGIN|" & nick & "|" & SeatOrDash(seat)
    Else
        mPendingLogin = True
        ConnectTo host
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim s As String
    s = Trim$(txtSay.Text)
    If Len(s) = 0 Then Exit Sub
    If sckTCP.State <> sckConnected Or Not mSignedIn Then
        AddSys IC_WARN, "Message not sent. You are not connected to the server.", C_RED
        Exit Sub
    End If
    SendLine "CHAT|" & s
    txtSay.Text = ""
    On Error Resume Next
    txtSay.SetFocus
End Sub

' The smiley button next to Send, drawn by hand like a normal button
Private Sub DrawSmiley(ByVal down As Boolean)
    Dim w As Long, h As Long, d As Long
    picSmiley.Cls
    w = picSmiley.ScaleWidth
    h = picSmiley.ScaleHeight
    If down Then
        Edge picSmiley, 0, 0, w, h, vb3DDKShadow, vb3DHighlight
        Edge picSmiley, 1, 1, w - 2, h - 2, vb3DShadow, vb3DLight
        d = 1
    Else
        Edge picSmiley, 0, 0, w, h, vb3DHighlight, vb3DDKShadow
        Edge picSmiley, 1, 1, w - 2, h - 2, vb3DLight, vb3DShadow
    End If
    DrawIcon16 picSmiley, IC_SMILEY, (w - 16) \ 2 + d, (h - 16) \ 2 + d
End Sub

Private Sub picSmiley_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then DrawSmiley True
End Sub

Private Sub picSmiley_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    DrawSmiley False
    If X >= 0 And Y >= 0 And X < picSmiley.ScaleWidth And Y < picSmiley.ScaleHeight Then PopupMenu mnuSmileys
End Sub

Private Sub mnuSmile_Click(Index As Integer)
    Dim s As String
    s = RTrim$(txtSay.Text)
    If Len(s) > 0 Then s = s & " "
    txtSay.Text = s & mnuSmile(Index).Caption & " "
    txtSay.SelStart = Len(txtSay.Text)
    On Error Resume Next
    txtSay.SetFocus
End Sub

Private Function PAChime() As String
    Dim i As Integer
    PAChime = "dingdong"
    For i = 0 To 3
        If optPA(i).Value Then
            Select Case i
            Case 0: PAChime = "dingdong"
            Case 1: PAChime = "triple"
            Case 2: PAChime = "horn"
            Case 3: PAChime = "none"
            End Select
        End If
    Next
End Function

Private Sub SendPA()
    Dim s As String
    s = Trim$(Replace(Replace(txtPA.Text, vbCr, " "), vbLf, " "))
    If Len(s) = 0 Then
        MsgBox "Enter the announcement text.", vbExclamation, "Announce"
        On Error Resume Next
        txtPA.SetFocus
        Exit Sub
    End If
    SendLine "PA|" & PAChime() & "|" & s
    txtPA.Text = ""
    Note "Announcement sent."
End Sub

Private Sub cmdPASend_Click()
    SendPA
End Sub

Private Sub cmdPATest_Click()
    PlayChime PAChime()
End Sub

Private Sub mnuFileCopy_Click()
    CopyChat
End Sub

Private Sub mnuFileSave_Click()
    SaveChat
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuGo_Click(Index As Integer)
    GoPage Index
End Sub

Private Sub mnuToolsSounds_Click()
    ToggleSounds
End Sub

Private Sub mnuToolsNotify_Click()
    frmNotify.Show vbModal, Me
End Sub

Private Sub mnuToolsOptions_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub mnuToolsCart_Click()
    ShowCartWizard
End Sub

Private Sub mnuFilePrint_Click()
    PrintSchedule
End Sub

Private Sub mnuEvIn_Click()
    ToggleIn mSchedSel
End Sub

Private Sub mnuEvRemind_Click()
    ToggleRemind mSchedSel
End Sub

Private Sub mnuEvInfo_Click()
    ShowEventInfo mSchedSel
End Sub

Private Sub mnuEvChat_Click()
    EventToChat mSchedSel
End Sub

Private Sub mnuEvDel_Click()
    RemoveSelectedEvent
End Sub

Private Sub mnuToolsHost_Click()
    Dim pw As String
    pw = InputBox("Host password:", "Host Mode")
    If Len(pw) > 0 Then SendLine "ADMIN|" & pw
End Sub

Private Sub mnuAdminPA_Click()
    GoPage PG_PA
End Sub

Private Sub mnuHelpPerf_Click()
    frmPerf.Show 0, Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

' For the About box
' The version set in Project > Properties > Make. The server compares it with the PalmsLAN.exe on its
' file page, so raise it for each release: PCs only update to a higher number.
Public Function AppVersion() As String
    AppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function ServerText() As String
    If Len(mServerName) > 0 And sckTCP.State = sckConnected Then
        ServerText = "Connected to " & mServerName & " at " & mServerHost & " (Palms Server " & mServerVer & ")."
    Else
        ServerText = "Not connected."
    End If
End Function

Private Sub UpdateMenus()
    Dim k As Integer
    mnuFileCopy.Enabled = mSignedIn
    mnuFileSave.Enabled = mSignedIn
    mnuFilePrint.Enabled = mSignedIn
    For k = 0 To PG_COUNT - 1
        mnuGo(k).Enabled = mSignedIn
    Next
    mnuGo(PG_PA).Visible = mIsAdmin
    mnuToolsHost.Enabled = mSignedIn And Not mIsAdmin
    mnuToolsHost.Checked = mSignedIn And mIsAdmin
    mnuAdmin.Visible = mSignedIn And mIsAdmin
    mnuAdminPA.Enabled = mSignedIn And mIsAdmin
    mnuAdminSched.Enabled = mSignedIn And mIsAdmin
    mnuAdminTpl.Enabled = mSignedIn And mIsAdmin
    mnuAdminMap.Enabled = mSignedIn And mIsAdmin
    mnuAdminVote.Enabled = mSignedIn And mIsAdmin And (mVoteState <> "open")
    mnuAdminVoteStop.Enabled = mSignedIn And mIsAdmin And (mVoteState = "open")
    mnuAdminCup.Enabled = mSignedIn And mIsAdmin And ServerHasCups()
    mnuAdminProj.Enabled = mSignedIn And mIsAdmin And ServerHasProjector()
    mnuViewProj.Enabled = mSignedIn
End Sub

Private Sub mnuViewProj_Click()
    frmProjector.Show
End Sub

Private Sub mnuAdminProj_Click()
    frmProjSet.Show vbModal, Me
End Sub

Private Sub mnuAdminSched_Click()
    frmEvent.Show vbModal, Me
End Sub

Private Sub mnuAdminTpl_Click()
    frmTemplates.Show vbModal, Me
End Sub

Private Sub mnuAdminMap_Click()
    ShowMapEditor
End Sub

Private Sub mnuAdminVote_Click()
    frmVote.Show vbModal, Me
End Sub

Private Sub mnuAdminVoteStop_Click()
    SendLine "VOTESTOP"
End Sub

Private Sub mnuAdminCup_Click()
    CupAction "cupnew"
End Sub

' Asked for lots of times in a row (people joining, taking breaks...): sent once, on the next tick
Private Sub RequestWho()
    mWantWho = True
    Want 0
End Sub

' For the seat window
Public Function MySeat() As String
    MySeat = mMySeat
End Function

Public Function IsOnline() As Boolean
    IsOnline = mSignedIn And (sckTCP.State = sckConnected)
End Function

' Servers from 0.4.1 know about seats, WHOIS and SPECS
Public Function ServerIsNew() As Boolean
    ServerIsNew = (VerNum(mServerVer) >= 401)
End Function

' "0.4.1" becomes 401, "0.4" becomes 400
Private Function VerNum(ByVal v As String) As Long
    Dim p() As String, i As Integer, n As Long
    p = Split(v & ".0.0", ".")
    For i = 0 To 2
        n = n * 100 + Val(p(i))
    Next
    VerNum = n
End Function

' ================================================================ small helpers

Private Function Fld(f() As String, ByVal i As Integer) As String
    If i <= UBound(f) Then Fld = f(i)
End Function

' Field i and everything after it, so messages can contain "|"
Private Function RestOf(f() As String, ByVal i As Integer) As String
    Dim k As Integer, s As String
    If i > UBound(f) Then Exit Function
    s = f(i)
    For k = i + 1 To UBound(f)
        s = s & "|" & f(k)
    Next
    RestOf = s
End Function

Private Function SeatText(ByVal seat As String) As String
    If seat <> "-" And Len(seat) > 0 Then SeatText = " at seat " & seat
End Function

Private Function SeatOrDash(ByVal seat As String) As String
    If Len(seat) = 0 Then
        SeatOrDash = "-"
    Else
        SeatOrDash = seat
    End If
End Function

Private Function ValidNick(ByVal s As String) As Boolean
    Dim i As Integer
    If Len(s) < 2 Or Len(s) > 12 Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "[A-Za-z0-9_-]") Then Exit Function
    Next
    ValidNick = True
End Function

Private Function ValidSeat(ByVal s As String) As Boolean
    ValidSeat = (s Like "[A-Z][0-9][0-9]") Or (s Like "[A-Z][A-Z][0-9][0-9]")
End Function
