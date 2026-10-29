VERSION 5.00
Begin VB.Form frmCartWizard
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Cartridge Wizard"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "frmCartWizard"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picWhite
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      Begin VB.PictureBox picWater
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5865
         Left            =   0
         ScaleHeight     =   391
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   164
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   2460
      End
      Begin VB.Frame fraPage
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   5400
         Index           =   0
         Left            =   2700
         TabIndex        =   2
         Top             =   240
         Width           =   6060
         Begin VB.Label lblWTitle
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the Make Cartridge Wizard"
            Height          =   780
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   6060
         End
         Begin VB.Label lblWText
            BackStyle       =   0  'Transparent
            Caption         =   "This wizard makes a cartridge for the Game Library from a game installed on this PC or on the network. The cartridge is uploaded to the Cartridge Server, and the host approves it before it appears in the Game Library."
            Height          =   1080
            Left            =   0
            TabIndex        =   4
            Top             =   960
            Width           =   6060
         End
         Begin VB.Label lblWAsk
            BackStyle       =   0  'Transparent
            Caption         =   "What do you want to do?"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   2220
            Width           =   6060
         End
         Begin VB.OptionButton optNew
            BackColor       =   &H80000005&
            Caption         =   "&Make a new cartridge"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   2550
            Value           =   -1  'True
            Width           =   5820
         End
         Begin VB.OptionButton optRebuild
            BackColor       =   &H80000005&
            Caption         =   "&Rebuild a cartridge (for example, after patching the game)"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   2910
            Width           =   5820
         End
         Begin VB.Label lblW7za
            BackStyle       =   0  'Transparent
            Caption         =   ""
            ForeColor       =   &H00101CAD&
            Height          =   675
            Left            =   0
            TabIndex        =   8
            Top             =   3480
            Width           =   6060
         End
         Begin VB.Label lblWNext
            BackStyle       =   0  'Transparent
            Caption         =   "To continue, click Next."
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   5145
            Width           =   6060
         End
      End
      Begin VB.Frame fraPage
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   5400
         Index           =   7
         Left            =   2700
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   6060
         Begin VB.Label lblDTitle
            BackStyle       =   0  'Transparent
            Caption         =   "Completing the Make Cartridge Wizard"
            Height          =   780
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   6060
         End
         Begin VB.Label lblDText
            BackStyle       =   0  'Transparent
            Caption         =   ""
            Height          =   2250
            Left            =   0
            TabIndex        =   12
            Top             =   960
            Width           =   6060
         End
         Begin VB.CommandButton cmdApprove
            Caption         =   "&Approve Cartridge..."
            Height          =   345
            Left            =   0
            TabIndex        =   13
            Top             =   3360
            Width           =   2400
         End
         Begin VB.CommandButton cmdOpenFolder
            Caption         =   "&Open Folder"
            Height          =   345
            Left            =   0
            TabIndex        =   14
            Top             =   3360
            Width           =   2400
         End
         Begin VB.CommandButton cmdAnother
            Caption         =   "Make &Another Cartridge"
            Height          =   345
            Left            =   2520
            TabIndex        =   15
            Top             =   3360
            Width           =   2400
         End
         Begin VB.Label lblDNext
            BackStyle       =   0  'Transparent
            Caption         =   "To close this wizard, click Finish."
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   5145
            Width           =   6060
         End
      End
   End
   Begin VB.PictureBox picHead
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Label lblSrcText
         Caption         =   "Select the folder the game is installed in, on this PC or on the network. Install the game first, and apply any patches and default settings."
         Height          =   450
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   8280
      End
      Begin VB.Label lblSrc
         Caption         =   "&Folder:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox txtSrc
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   630
         Width           =   6075
      End
      Begin VB.CommandButton cmdSrcBrowse
         Caption         =   "Browse..."
         Height          =   345
         Left            =   7155
         TabIndex        =   22
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label lblScan
         Caption         =   ""
         Height          =   675
         Left            =   960
         TabIndex        =   23
         Top             =   1110
         Width           =   7320
      End
      Begin VB.Label lblPrev
         Caption         =   ""
         Height          =   900
         Left            =   960
         TabIndex        =   24
         Top             =   1860
         Width           =   7320
      End
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Label lblName
         Caption         =   "N&ame:"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   60
         Width           =   1260
      End
      Begin VB.TextBox txtName
         Height          =   315
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   27
         Top             =   0
         Width           =   4500
      End
      Begin VB.Label lblNameHint
         Caption         =   "As shown in the Game Library, for example Unreal Tournament"
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   1320
         TabIndex        =   28
         Top             =   360
         Width           =   6960
      End
      Begin VB.Label lblId
         Caption         =   "&ID:"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   1260
      End
      Begin VB.TextBox txtId
         Height          =   315
         Left            =   1320
         MaxLength       =   24
         TabIndex        =   30
         Top             =   660
         Width           =   2100
      End
      Begin VB.Label lblIdHint
         Caption         =   "Lowercase letters, numbers and -, for example ut99"
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   3540
         TabIndex        =   31
         Top             =   720
         Width           =   4740
      End
      Begin VB.Label lblIdNote
         Caption         =   "Each game needs a unique ID. A newer cartridge with the same ID replaces the older one."
         ForeColor       =   &H00636163&
         Height          =   420
         Left            =   1320
         TabIndex        =   32
         Top             =   1020
         Width           =   6960
      End
      Begin VB.Label lblVer
         Caption         =   "&Version:"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   1560
         Width           =   1260
      End
      Begin VB.TextBox txtVer
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   34
         Top             =   1500
         Width           =   2100
      End
      Begin VB.Label lblVerHint
         Caption         =   "Optional, for example v436 or 1.1.1.1"
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   3540
         TabIndex        =   35
         Top             =   1560
         Width           =   4740
      End
      Begin VB.Label lblRev
         Caption         =   "&Revision:"
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   1980
         Width           =   1260
      End
      Begin VB.TextBox txtRev
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label lblRevHint
         Caption         =   "Goes up by 1 each time this cartridge is rebuilt"
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   2040
         TabIndex        =   38
         Top             =   1980
         Width           =   6240
      End
      Begin VB.Label lblPlayers
         Caption         =   "&Players:"
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   2400
         Width           =   1260
      End
      Begin VB.TextBox txtPlayers
         Height          =   315
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   40
         Top             =   2340
         Width           =   900
      End
      Begin VB.Label lblPlayersHint
         Caption         =   "For example, 16 or 2-16"
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   2340
         TabIndex        =   41
         Top             =   2400
         Width           =   5940
      End
      Begin VB.Label lblFolder
         Caption         =   "Install &folder:"
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Top             =   3060
         Width           =   1260
      End
      Begin VB.Label lblRoot
         Caption         =   "C:\LANPartyGames\"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox txtFolder
         Height          =   315
         Left            =   2850
         MaxLength       =   40
         TabIndex        =   44
         Top             =   3000
         Width           =   2700
      End
      Begin VB.Label lblFolderHint
         Caption         =   "Letters, numbers, -, _ and . only. Each PC installs the game in this folder, inside its games folder."
         ForeColor       =   &H00636163&
         Height          =   450
         Left            =   1320
         TabIndex        =   45
         Top             =   3390
         Width           =   6960
      End
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   3
      Left            =   360
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.OptionButton optSingle
         Caption         =   "&Single game"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.OptionButton optMulti
         Caption         =   "&Multicart (several games on one cartridge)"
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   0
         Width           =   4500
      End
      Begin VB.Label lblMultiHint
         Caption         =   ""
         ForeColor       =   &H00636163&
         Height          =   225
         Left            =   6480
         TabIndex        =   49
         Top             =   15
         Width           =   1800
      End
      Begin VB.Frame fraSingle
         BorderStyle     =   0  'None
         Height          =   4260
         Left            =   0
         TabIndex        =   50
         Top             =   390
         Width           =   8280
         Begin VB.Label lblExe
            Caption         =   "P&rogram:"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   60
            Width           =   1380
         End
         Begin VB.ComboBox cboExe
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   0
            Width           =   5595
         End
         Begin VB.CommandButton cmdExeBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   53
            Top             =   -15
            Width           =   1125
         End
         Begin VB.Label lblArgs
            Caption         =   "P&arameters:"
            Height          =   255
            Left            =   0
            TabIndex        =   54
            Top             =   480
            Width           =   1380
         End
         Begin VB.TextBox txtArgs
            Height          =   315
            Left            =   1440
            TabIndex        =   55
            Top             =   420
            Width           =   6840
         End
         Begin VB.Label lblArgsHint
            Caption         =   "Optional, for example -nosound"
            ForeColor       =   &H00636163&
            Height          =   225
            Left            =   1440
            TabIndex        =   56
            Top             =   780
            Width           =   6840
         End
         Begin VB.Label lblCfg
            Caption         =   "&Configuration:"
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   1140
            Width           =   1380
         End
         Begin VB.ComboBox cboCfg
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1080
            Width           =   5595
         End
         Begin VB.CommandButton cmdCfgBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   59
            Top             =   1065
            Width           =   1125
         End
         Begin VB.Label lblCfgArgs
            Caption         =   "Paramet&ers:"
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   1560
            Width           =   1380
         End
         Begin VB.TextBox txtCfgArgs
            Height          =   315
            Left            =   1440
            TabIndex        =   61
            Top             =   1500
            Width           =   6840
         End
         Begin VB.Label lblIcon
            Caption         =   "&Icon:"
            Height          =   255
            Left            =   0
            TabIndex        =   62
            Top             =   1980
            Width           =   1380
         End
         Begin VB.TextBox txtIcon
            Height          =   315
            Left            =   1440
            TabIndex        =   63
            Top             =   1920
            Width           =   5595
         End
         Begin VB.CommandButton cmdIconBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   64
            Top             =   1905
            Width           =   1125
         End
         Begin VB.Label lblIconHint
            Caption         =   "An .ico, .bmp, .gif or .jpg file, or a program that contains an icon. Leave blank to use the program's icon."
            ForeColor       =   &H00636163&
            Height          =   450
            Left            =   1440
            TabIndex        =   65
            Top             =   2280
            Width           =   6840
         End
      End
      Begin VB.Frame fraMulti
         BorderStyle     =   0  'None
         Height          =   4260
         Left            =   0
         TabIndex        =   66
         Top             =   390
         Visible         =   0   'False
         Width           =   8280
         Begin VB.Label lblPackArgs
            Caption         =   "&All games:"
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   60
            Width           =   1260
         End
         Begin VB.TextBox txtPackArgs
            Height          =   315
            Left            =   1320
            TabIndex        =   68
            ToolTipText     =   "Parameters added after each game's own"
            Top             =   0
            Width           =   6960
         End
         Begin VB.ListBox lstGames
            Height          =   3210
            Left            =   0
            Style           =   1  'Checkbox
            TabIndex        =   69
            Top             =   450
            Width           =   2640
         End
         Begin VB.CommandButton cmdAddGame
            Caption         =   "A&dd Game"
            Height          =   345
            Left            =   0
            TabIndex        =   70
            Top             =   3780
            Width           =   1260
         End
         Begin VB.CommandButton cmdRemoveGame
            Caption         =   "Remo&ve"
            Height          =   345
            Left            =   1380
            TabIndex        =   71
            Top             =   3780
            Width           =   1260
         End
         Begin VB.Label lblGCmd
            Caption         =   ""
            ForeColor       =   &H00636163&
            Height          =   420
            Left            =   2820
            TabIndex        =   72
            Top             =   450
            Width           =   5460
         End
         Begin VB.Label lblGName
            Caption         =   "Nam&e:"
            Height          =   255
            Left            =   2820
            TabIndex        =   73
            Top             =   960
            Width           =   1020
         End
         Begin VB.TextBox txtGName
            Height          =   315
            Left            =   3900
            MaxLength       =   40
            TabIndex        =   74
            Top             =   900
            Width           =   4380
         End
         Begin VB.Label lblGPlayers
            Caption         =   "Pla&yers:"
            Height          =   255
            Left            =   2820
            TabIndex        =   75
            Top             =   1380
            Width           =   1020
         End
         Begin VB.TextBox txtGPlayers
            Height          =   315
            Left            =   3900
            MaxLength       =   7
            TabIndex        =   76
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label lblGPlayersHint
            Caption         =   ""
            ForeColor       =   &H00636163&
            Height          =   225
            Left            =   4590
            TabIndex        =   77
            Top             =   1380
            Width           =   3690
         End
         Begin VB.Label lblGExe
            Caption         =   "&Program:"
            Height          =   255
            Left            =   2820
            TabIndex        =   78
            Top             =   1800
            Width           =   1020
         End
         Begin VB.ComboBox cboGExe
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1740
            Width           =   3195
         End
         Begin VB.CommandButton cmdGExeBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   80
            Top             =   1725
            Width           =   1125
         End
         Begin VB.Label lblGArgs
            Caption         =   "Parame&ters:"
            Height          =   255
            Left            =   2820
            TabIndex        =   81
            Top             =   2220
            Width           =   1020
         End
         Begin VB.TextBox txtGArgs
            Height          =   315
            Left            =   3900
            TabIndex        =   82
            Top             =   2160
            Width           =   4380
         End
         Begin VB.Label lblGStart
            Caption         =   "Sta&rt in:"
            Height          =   255
            Left            =   2820
            TabIndex        =   83
            Top             =   2640
            Width           =   1020
         End
         Begin VB.TextBox txtGStart
            Height          =   315
            Left            =   3900
            TabIndex        =   84
            ToolTipText     =   "A folder inside the game folder. Leave blank to use the program's folder."
            Top             =   2580
            Width           =   3195
         End
         Begin VB.CommandButton cmdGStartBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   85
            Top             =   2565
            Width           =   1125
         End
         Begin VB.Label lblGIcon
            Caption         =   "&Icon:"
            Height          =   255
            Left            =   2820
            TabIndex        =   86
            Top             =   3060
            Width           =   1020
         End
         Begin VB.TextBox txtGIcon
            Height          =   315
            Left            =   3900
            TabIndex        =   87
            Top             =   3000
            Width           =   3195
         End
         Begin VB.CommandButton cmdGIconBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   88
            Top             =   2985
            Width           =   1125
         End
         Begin VB.Label lblGCfg
            Caption         =   "Con&fig:"
            Height          =   255
            Left            =   2820
            TabIndex        =   89
            Top             =   3480
            Width           =   1020
         End
         Begin VB.ComboBox cboGCfg
            Height          =   315
            Left            =   3900
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   3420
            Width           =   3195
         End
         Begin VB.CommandButton cmdGCfgBrowse
            Caption         =   "Browse..."
            Height          =   345
            Left            =   7155
            TabIndex        =   91
            Top             =   3405
            Width           =   1125
         End
         Begin VB.Label lblGCfgArgs
            Caption         =   "Parameters:"
            Height          =   255
            Left            =   2820
            TabIndex        =   92
            Top             =   3900
            Width           =   1020
         End
         Begin VB.TextBox txtGCfgArgs
            Height          =   315
            Left            =   3900
            TabIndex        =   93
            Top             =   3840
            Width           =   4380
         End
      End
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   4
      Left            =   360
      TabIndex        =   94
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Label lblRegText
         Caption         =   "Only for games that need registry settings to run (Half-Life, for example). On a PC with the game installed, run regedit, select the game's key and click File > Export."
         Height          =   450
         Left            =   0
         TabIndex        =   95
         Top             =   0
         Width           =   8280
      End
      Begin VB.Label lblReg
         Caption         =   "Registry &file:"
         Height          =   255
         Left            =   0
         TabIndex        =   96
         Top             =   630
         Width           =   1440
      End
      Begin VB.ComboBox cboReg
         Height          =   315
         Left            =   1560
         Style           =   0  'Dropdown Combo
         TabIndex        =   97
         Top             =   570
         Width           =   5475
      End
      Begin VB.CommandButton cmdRegBrowse
         Caption         =   "Browse..."
         Height          =   345
         Left            =   7155
         TabIndex        =   98
         Top             =   555
         Width           =   1125
      End
      Begin VB.Label lblRegInfo
         Caption         =   ""
         Height          =   840
         Left            =   1560
         TabIndex        =   99
         Top             =   960
         Width           =   6720
      End
      Begin VB.CheckBox chkRegAnyway
         Caption         =   "&Use this file anyway"
         Height          =   255
         Left            =   1560
         TabIndex        =   100
         Top             =   1830
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Label lblRegPath
         Caption         =   "&Install path:"
         Height          =   255
         Left            =   0
         TabIndex        =   101
         Top             =   2250
         Width           =   1440
      End
      Begin VB.TextBox txtRegPath
         Height          =   315
         Left            =   1560
         TabIndex        =   102
         Top             =   2190
         Width           =   6720
      End
      Begin VB.Label lblRegPathHint
         Caption         =   "The install folder used in the registry file. On each PC it is replaced with the folder the game is installed in."
         ForeColor       =   &H00636163&
         Height          =   420
         Left            =   1560
         TabIndex        =   103
         Top             =   2550
         Width           =   6720
      End
      Begin VB.Label lblNotesText
         Caption         =   "Setup notes: a text file shown to players after installation, for steps each PC needs (for example, entering a CD key)."
         Height          =   450
         Left            =   0
         TabIndex        =   104
         Top             =   3150
         Width           =   8280
      End
      Begin VB.Label lblNotes
         Caption         =   "Se&tup notes:"
         Height          =   255
         Left            =   0
         TabIndex        =   105
         Top             =   3810
         Width           =   1440
      End
      Begin VB.TextBox txtNotes
         Height          =   315
         Left            =   1560
         TabIndex        =   106
         Top             =   3750
         Width           =   4275
      End
      Begin VB.CommandButton cmdNotesBrowse
         Caption         =   "Browse..."
         Height          =   345
         Left            =   5955
         TabIndex        =   107
         Top             =   3735
         Width           =   1125
      End
      Begin VB.CommandButton cmdNotesView
         Caption         =   "&View"
         Height          =   345
         Left            =   7155
         TabIndex        =   108
         Top             =   3735
         Width           =   1125
      End
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   5
      Left            =   360
      TabIndex        =   109
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Label lblOut
         Caption         =   "&Save in:"
         Height          =   255
         Left            =   0
         TabIndex        =   110
         Top             =   60
         Width           =   960
      End
      Begin VB.TextBox txtOut
         Height          =   315
         Left            =   1080
         TabIndex        =   111
         Top             =   0
         Width           =   5955
      End
      Begin VB.CommandButton cmdOutBrowse
         Caption         =   "Browse..."
         Height          =   345
         Left            =   7155
         TabIndex        =   112
         Top             =   -15
         Width           =   1125
      End
      Begin VB.Label lblOutHint
         Caption         =   ""
         ForeColor       =   &H00636163&
         Height          =   420
         Left            =   1080
         TabIndex        =   113
         Top             =   360
         Width           =   7200
      End
      Begin VB.PictureBox picSum
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   2520
         Left            =   0
         ScaleHeight     =   164
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   548
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   840
         Width           =   8280
      End
      Begin VB.Label lblProblems
         Caption         =   ""
         Height          =   1170
         Left            =   0
         TabIndex        =   115
         Top             =   3480
         Width           =   8280
      End
   End
   Begin VB.Frame fraPage
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   6
      Left            =   360
      TabIndex        =   116
      Top             =   1080
      Visible         =   0   'False
      Width           =   8280
      Begin VB.Label lblMaking
         Caption         =   ""
         Height          =   255
         Left            =   0
         TabIndex        =   117
         Top             =   240
         Width           =   8280
      End
      Begin VB.PictureBox picProg
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   552
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   660
         Width           =   8280
      End
      Begin VB.Label lblMakingSub
         Caption         =   ""
         Height          =   450
         Left            =   0
         TabIndex        =   119
         Top             =   1080
         Width           =   8280
      End
      Begin VB.Label lblMakingNote
         Caption         =   "Palms LAN Center can be used while the cartridge is made. Large games take several minutes. To cancel, click Stop."
         ForeColor       =   &H00636163&
         Height          =   675
         Left            =   0
         TabIndex        =   120
         Top             =   1860
         Width           =   8280
      End
   End
   Begin VB.CommandButton cmdBack
      Caption         =   "< &Back"
      Height          =   345
      Left            =   5325
      TabIndex        =   121
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   345
      Left            =   6450
      TabIndex        =   122
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Cancel          =   -1  'True
      Height          =   345
      Left            =   7695
      TabIndex        =   123
      Top             =   6030
      Width           =   1125
   End
   Begin VB.Timer tmrWork
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   6000
   End
   Begin VB.Line linSep2
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   9000
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line linSep1
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   9000
      Y1              =   5865
      Y2              =   5865
   End
End
Attribute VB_Name = "frmCartWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Palms LAN Center - The Make Cartridge Wizard (the work is done by modCartridge).
Option Explicit

Private Const PG_WELCOME As Integer = 0
Private Const PG_FOLDER As Integer = 1
Private Const PG_DETAILS As Integer = 2
Private Const PG_GAMES As Integer = 3
Private Const PG_EXTRAS As Integer = 4
Private Const PG_READY As Integer = 5
Private Const PG_MAKING As Integer = 6
Private Const PG_DONE As Integer = 7
Private Const MAX_GAMES As Long = 60
Private Const REG_PATH_HINT As String = "The install folder used in the registry file. On each PC it is replaced with the folder the game is installed in."

Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private mPage As Integer
Private mFilling As Boolean          ' boxes being filled in by the program, not typed in
Private mJob As String               ' what tmrWork is doing: scan, server or build
Private mScanThenNext As Boolean
Private mFocus As Object             ' the box with the first problem on a page

' ---- the game folder
Private mSrc As String
Private mScanned As Boolean
Private mScanning As Boolean
Private mFiles As Long
Private mSizeText As String
Private mPieces As Long
Private mExes() As String
Private mExeCount As Long
Private mMore As String              ' programs picked with Browse, deeper in the folder, separated by |
Private mRegsFound() As String
Private mRegCount As Long

' ---- details
Private mName As String
Private mCartId As String
Private mCartIdTouched As Boolean
Private mVer As String
Private mRev As String          ' the cartridge's revision: 1, then 2 when it's rebuilt...
Private mPlayers As String
Private mFolder As String
Private mFolderTouched As Boolean

' ---- the games: one, or several on a multicart
Private mPack As Boolean
Private mCands As Long               ' games found in the folder (or in the previous build)
Private mExe As String
Private mArgs As String
Private mCfg As String
Private mCfgArgs As String
Private mIcon As String
Private mPackArgs As String
Private mGOn(0 To 59) As Boolean
Private mGName(0 To 59) As String
Private mGPlayers(0 To 59) As String
Private mGExe(0 To 59) As String
Private mGArgs(0 To 59) As String
Private mGStart(0 To 59) As String
Private mGIcon(0 To 59) As String
Private mGCfg(0 To 59) As String
Private mGCfgArgs(0 To 59) As String
Private mGCount As Long
Private mSel As Long

' ---- registry file and setup notes
Private mReg As String
Private mRegFull As String
Private mRegKeys As Long
Private mRegValues As Long
Private mRegNotes As String
Private mRegTrouble As String
Private mRegPath As String
Private mRegAnyway As Boolean
Private mNotes As String

' ---- saving and making
Private mOut As String
Private mOutChecked As Boolean
Private mReadyOK As Boolean
Private mPrevFile As String
Private mPrevApplied As Boolean
Private mBuilding As Boolean
Private mStopping As Boolean
Private mBuilt As Boolean
Private mBuildStart As Long
Private mBuildSecs As Long
Private mPct As Long
Private mResult As String            ' from EngBuildPoll: done|..., fail|..., stopped

' ================================================================ opening and closing

Private Sub Form_Load()
    lblWTitle.FontSize = 12
    lblWTitle.FontBold = True
    lblDTitle.FontSize = 12
    lblDTitle.FontBold = True
    lblMaking.FontBold = True
    lblRoot.Caption = "C:\LANPartyGames\"
    DrawWater
    mOut = ""
    StartOver
End Sub

' Making a cartridge right now (the main window asks before closing)
Public Function IsBuilding() As Boolean
    IsBuilding = mBuilding
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If Not ConfirmClose() Then
            Cancel = 1
            Exit Sub
        End If
    End If
    ' closing while 7-Zip is packing: stop it, and take away the half-made cartridge
    If mBuilding Then
        tmrWork.Enabled = False
        EngBuildAbandon
        mBuilding = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrWork.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Function ConfirmClose() As Boolean
    If mBuilding Then
        ConfirmClose = (MsgBox("Stop making the cartridge? The incomplete cartridge will be removed.", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes)
    ElseIf mScanned And Not mBuilt Then
        ConfirmClose = (MsgBox("Close the Make Cartridge Wizard? The settings entered will be lost.", vbQuestion + vbYesNo + vbDefaultButton2, Caption) = vbYes)
    Else
        ConfirmClose = True
    End If
End Function

' Everything back to the start, apart from where cartridges are saved
Private Sub StartOver()
    ResetAnswers
    ShowPage PG_WELCOME
End Sub

Private Sub ResetAnswers()
    EngInit AppDir(), frmMain.AppVersion()
    If Len(CartServerHost()) > 0 Then EngSetServer CartServerHost()
    mSrc = ""
    mScanned = False
    mScanning = False
    mFiles = 0
    mSizeText = ""
    mPieces = 1
    mExeCount = 0
    ReDim mExes(0 To 0)
    mRegCount = 0
    ReDim mRegsFound(0 To 0)
    mMore = ""
    mName = ""
    mCartId = ""
    mCartIdTouched = False
    mVer = ""
    mRev = "1"
    mPlayers = ""
    mFolder = ""
    mFolderTouched = False
    mPack = False
    mCands = 0
    mExe = ""
    mArgs = ""
    mCfg = ""
    mCfgArgs = ""
    mIcon = ""
    mPackArgs = ""
    mGCount = 0
    mSel = -1
    mReg = ""
    mRegFull = ""
    mRegPath = ""
    mRegAnyway = False
    mNotes = ""
    mPrevFile = ""
    mPrevApplied = False
    mBuilt = False
    mBuilding = False
    mStopping = False
    mResult = ""
    mFilling = True
    txtSrc.Text = ""
    optNew.Value = True
    mFilling = False
    lblScan.Caption = ""
    lblPrev.Caption = ""
    If Eng7zaThere() Then
        lblW7za.Caption = ""
    Else
        lblW7za.Caption = "7za.exe is missing from " & EngHere() & ". It is needed to make cartridges. Run palms-update.vbs to install it."
    End If
End Sub

' The Cartridge Server, from the address the Palms Server gives out (like http://10.0.200.73/)
Private Function CartServerHost() As String
    Dim s As String, p As Long
    s = gPakBase
    p = InStr(s, "://")
    If p > 0 Then s = Mid$(s, p + 3)
    p = InStr(s, "/")
    If p > 0 Then s = Left$(s, p - 1)
    p = InStr(s, ":")
    If p > 0 Then s = Left$(s, p - 1)
    CartServerHost = s
End Function

' ================================================================ moving between pages

Private Sub ShowPage(ByVal n As Integer)
    Dim i As Integer
    mPage = n
    picWhite.Visible = (n = PG_WELCOME Or n = PG_DONE)
    picHead.Visible = Not picWhite.Visible
    For i = 0 To 7
        fraPage(i).Visible = (i = n)
    Next
    Select Case n
    Case PG_FOLDER
        FillFolder
    Case PG_DETAILS
        FillDetails
    Case PG_GAMES
        FillGames
    Case PG_EXTRAS
        FillExtras
    Case PG_READY
        FillReady
    Case PG_DONE
        FillDone
    End Select
    DrawHead
    UpdateButtons
    FocusFirst
End Sub

Private Sub UpdateButtons()
    cmdBack.Enabled = mPage > PG_WELCOME And mPage <> PG_MAKING And Not mScanning And Not (mPage = PG_DONE And mBuilt)
    Select Case mPage
    Case PG_READY
        cmdNext.Caption = "&Make"
    Case PG_DONE
        cmdNext.Caption = "Finish"
    Case Else
        cmdNext.Caption = "&Next >"
    End Select
    cmdNext.Enabled = Not mScanning And mPage <> PG_MAKING And Not (mPage = PG_READY And Not mReadyOK)
    If mPage = PG_MAKING Then
        cmdCancel.Caption = "&Stop"
        cmdCancel.Enabled = Not mStopping
    Else
        cmdCancel.Caption = "Cancel"
        cmdCancel.Enabled = True
    End If
    cmdCancel.Visible = (mPage <> PG_DONE)
End Sub

Private Sub FocusFirst()
    On Error Resume Next
    Select Case mPage
    Case PG_WELCOME
        If optRebuild.Value Then optRebuild.SetFocus Else optNew.SetFocus
    Case PG_FOLDER
        txtSrc.SetFocus
    Case PG_DETAILS
        txtName.SetFocus
    Case PG_GAMES
        If mPack Then lstGames.SetFocus Else cboExe.SetFocus
    Case PG_EXTRAS
        cboReg.SetFocus
    Case PG_READY
        txtOut.SetFocus
    Case PG_MAKING
        cmdCancel.SetFocus
    Case Else
        cmdNext.SetFocus
    End Select
End Sub

Private Sub cmdNext_Click()
    Select Case mPage
    Case PG_WELCOME
        If optRebuild.Value Then
            OpenPrevious
        Else
            ' a new one after loading a cartridge to rebuild: none of its settings
            If Len(mPrevFile) > 0 Then ResetAnswers
            ShowPage PG_FOLDER
        End If
    Case PG_FOLDER
        If Not mScanned Or StrComp(Trim$(txtSrc.Text), mSrc, vbTextCompare) <> 0 Then
            StartScan True
        Else
            ShowPage PG_DETAILS
        End If
    Case PG_DETAILS
        Collect PG_DETAILS
        If PageOK(PG_DETAILS) Then ShowPage PG_GAMES
    Case PG_GAMES
        Collect PG_GAMES
        If PageOK(PG_GAMES) Then ShowPage PG_EXTRAS
    Case PG_EXTRAS
        Collect PG_EXTRAS
        If PageOK(PG_EXTRAS) Then ShowPage PG_READY
    Case PG_READY
        Collect PG_READY
        Make
    Case PG_DONE
        Unload Me
    End Select
End Sub

Private Sub cmdBack_Click()
    If mScanning Or mBuilding Then Exit Sub
    Select Case mPage
    Case PG_DONE
        ShowPage PG_READY
    Case PG_FOLDER
        ShowPage PG_WELCOME
    Case Else
        Collect mPage
        ShowPage mPage - 1
    End Select
End Sub

Private Sub cmdCancel_Click()
    If mPage = PG_MAKING Then
        StopBuild
        Exit Sub
    End If
    If ConfirmClose() Then Unload Me
End Sub

' The boxes on a page, into the answers
Private Sub Collect(ByVal n As Integer)
    Select Case n
    Case PG_DETAILS
        mName = Trim$(txtName.Text)
        mCartId = LCase$(Trim$(txtId.Text))
        mVer = Trim$(txtVer.Text)
        mRev = Trim$(txtRev.Text)
        mPlayers = Replace(Trim$(txtPlayers.Text), " ", "")
        mFolder = Trim$(txtFolder.Text)
    Case PG_GAMES
        If mPack Then
            mPackArgs = Trim$(txtPackArgs.Text)
            SaveGame
        Else
            mExe = ComboValue(cboExe)
            mArgs = Trim$(txtArgs.Text)
            mCfg = ComboValue(cboCfg)
            mCfgArgs = Trim$(txtCfgArgs.Text)
            mIcon = Trim$(txtIcon.Text)
        End If
    Case PG_EXTRAS
        If StrComp(Trim$(cboReg.Text), mReg, vbTextCompare) <> 0 Then RegChanged
        mRegPath = Trim$(txtRegPath.Text)
        mRegAnyway = (chkRegAnyway.Value = vbChecked)
        mNotes = Trim$(txtNotes.Text)
    Case PG_READY
        mOut = Trim$(txtOut.Text)
    End Select
End Sub

' ================================================================ drawing

' The white band at the top of each step, with the Palms sunset under it
Private Sub DrawHead()
    Dim t As String, s As String, ic As Integer, w As Long
    Select Case mPage
    Case PG_FOLDER
        t = "Game Folder"
        s = "Select the folder the game is installed in."
        ic = IC_FOLDER
    Case PG_DETAILS
        t = "Details"
        s = "How the game appears in the Game Library, and where it is installed."
        ic = IC_PENCIL
    Case PG_GAMES
        t = "Games"
        s = "The programs that start the game. A folder with several games can be made into a multicart."
        ic = IC_JOYSTICK
    Case PG_EXTRAS
        t = "Registry and Setup Notes"
        s = "Optional. Most games need neither."
        ic = IC_REG
    Case PG_READY
        t = "Ready to Make the Cartridge"
        s = "Check the details, then click Make."
        ic = IC_BOX
    Case PG_MAKING
        t = "Making the Cartridge"
        s = "7-Zip is packing the game folder."
        ic = IC_BOX
    Case Else
        Exit Sub
    End Select
    w = picHead.ScaleWidth
    picHead.Cls
    TextAt picHead, 16, 11, t, C_TITLE, True
    TextAt picHead, 32, 30, FitText(picHead, s, w - 32 - 64), C_SUB
    DrawIcon32 picHead, ic, w - 48, 13
    Sunset picHead, 0, 58, w, 3
End Sub

' The picture down the left of the first and last pages: Palms colours, a palm and a cartridge
Private Sub DrawWater()
    Dim i As Long, h As Long, w As Long, k As Double
    w = picWater.ScaleWidth
    h = picWater.ScaleHeight
    picWater.Cls
    For i = 0 To h - 1
        k = i / h
        Box picWater, 0, i, w, 1, RGB(49 + (8 - 49) * k, 97 + (36 - 97) * k, 107 + (66 - 107) * k)
    Next
    Box picWater, 0, h - 60, w, 6, RGB(247, 142, 33)
    Box picWater, 0, h - 54, w, 6, RGB(250, 160, 31)
    Box picWater, 0, h - 48, w, 6, RGB(253, 182, 30)
    Box picWater, 0, h - 42, w, 6, RGB(255, 205, 37)
    Box picWater, 0, h - 36, w, 6, RGB(255, 219, 41)
    DrawIcon32 picWater, IC_PALM, 14, 14
    TextAt picWater, 54, 23, "Palms Arcade", vbWhite, True
    DrawIcon32 picWater, IC_BOX, (w - 32) \ 2, h - 120
End Sub

' ================================================================ 1. the game folder

Private Sub FillFolder()
    If mScanned Then
        mFilling = True
        txtSrc.Text = mSrc
        mFilling = False
    End If
    RenderScan
End Sub

Private Sub cmdSrcBrowse_Click()
    Dim p As String
    p = EngPickFolder("Select the game folder:", "")
    If Len(p) = 0 Then Exit Sub
    txtSrc.Text = p
    StartScan False
End Sub

Private Sub StartScan(ByVal thenNext As Boolean)
    Dim p As String
    If mScanning Then Exit Sub
    p = Trim$(txtSrc.Text)
    If Len(p) = 0 Then
        MsgBox "Enter the game folder, or click Browse.", vbExclamation, Caption
        On Error Resume Next
        txtSrc.SetFocus
        Exit Sub
    End If
    If Not EngScanStart(p) Then
        MsgBox "Folder not found: " & p & vbCrLf & vbCrLf & "If it is on the network, check the connection.", vbExclamation, Caption
        Exit Sub
    End If
    mScanning = True
    mScanned = False
    mScanThenNext = thenNext
    lblScan.ForeColor = vbButtonText
    lblScan.Caption = "Scanning..."
    UpdateButtons
    Screen.MousePointer = vbArrowHourglass
    mJob = "scan"
    tmrWork.Interval = 10
    tmrWork.Enabled = True
End Sub

Private Sub tmrWork_Timer()
    Select Case mJob
    Case "scan"
        ScanTick
    Case "server"
        ServerTick
    Case "build"
        BuildTick
    Case Else
        tmrWork.Enabled = False
    End Select
End Sub

' A bit more of the folder, a quarter of a second at a time so the window keeps answering
Private Sub ScanTick()
    Dim done As Boolean
    tmrWork.Enabled = False
    done = EngScanStep()
    lblScan.Caption = "Scanning... " & NumText(EngScanInfo("files")) & " files (" & EngScanInfo("size") & ")"
    If Not done Then
        tmrWork.Enabled = True
        Exit Sub
    End If
    mScanning = False
    mJob = ""
    Screen.MousePointer = vbDefault
    Scanned
    UpdateButtons
    If mScanThenNext Then ShowPage PG_DETAILS
End Sub

' The folder's been looked through: work out the suggestions
Private Sub Scanned()
    Dim newSrc As Boolean, i As Long
    newSrc = (StrComp(EngScanInfo("src"), mSrc, vbTextCompare) <> 0)
    mSrc = EngScanInfo("src")
    mFilling = True
    txtSrc.Text = mSrc
    mFilling = False
    mScanned = True
    mFiles = EngScanInfo("files")
    mSizeText = EngScanInfo("size")
    mPieces = EngScanInfo("pieces")
    mExeCount = EngScanInfo("exes")
    ReDim mExes(0 To mExeCount)
    For i = 0 To mExeCount - 1
        mExes(i) = EngExe(i)
    Next
    mRegCount = EngScanInfo("regs")
    ReDim mRegsFound(0 To mRegCount)
    For i = 0 To mRegCount - 1
        mRegsFound(i) = EngRegFile(i)
    Next
    If Len(mPrevFile) > 0 And Not mPrevApplied Then
        ApplyPrevious
    ElseIf newSrc Then
        FreshGuesses
    ElseIf Not mPrevApplied Then
        ' the same folder again (shortcuts copied in since?): look for its games again, keep the details
        FindGames
    End If
    RenderScan
End Sub

Private Sub FreshGuesses()
    Dim nm As String
    nm = Mid$(mSrc, InStrRev(mSrc, "\") + 1)
    mName = EngTidy(nm)
    mCartIdTouched = False
    mFolderTouched = False
    mCartId = EngMakeId(mName)
    mFolder = EngMakeFolderName(mName)
    mVer = ""
    mRev = "1"
    mPlayers = ""
    mExe = EngGuessExe(mName, mCartId)
    mArgs = ""
    mCfg = ""
    mCfgArgs = ""
    mIcon = ""
    mMore = ""
    If mRegCount = 1 Then
        mReg = mRegsFound(0)
    Else
        mReg = ""
    End If
    mRegPath = ""
    mRegFull = ""
    mRegAnyway = False
    If Len(mReg) > 0 Then CheckReg True
    mNotes = ""
    FindGames
End Sub

' The games in the folder: from its shortcuts and mod folders
Private Sub FindGames()
    Dim i As Long, n As Long
    n = EngDetect()
    If n > MAX_GAMES Then n = MAX_GAMES
    mGCount = 0
    For i = 0 To n - 1
        AddGameRow True, EngCand(i, "name"), "", EngCand(i, "exe"), EngCand(i, "args"), EngCand(i, "start"), EngCand(i, "icon"), "", ""
    Next
    mCands = n
    mPackArgs = EngSharedArgs()
    mPack = (n >= 2)
    If mGCount > 0 Then
        mSel = 0
    Else
        mSel = -1
    End If
End Sub

Private Sub RenderScan()
    Dim s As String
    If mScanning Then Exit Sub
    If Not mScanned Then
        lblScan.Caption = ""
        Exit Sub
    End If
    lblScan.ForeColor = vbButtonText
    s = NumText(mFiles) & " files, " & mSizeText & ". "
    If mFiles = 0 Then
        s = "The folder is empty."
        lblScan.ForeColor = C_RED
    ElseIf mPack And mGCount > 0 And mPrevApplied Then
        s = s & "Previous build: " & mGCount & " games."
    ElseIf mPack And mGCount > 0 Then
        s = s & mGCount & " games found (from shortcuts and mod folders)."
    ElseIf mExeCount > 0 Then
        s = s & mExeCount & " program" & Plural(mExeCount) & " found."
    Else
        s = "No programs (.exe) were found in the top folders. Check that this is the correct folder."
        lblScan.ForeColor = C_RED
    End If
    lblScan.Caption = s & " Click Next to continue."
End Sub

' ---- rebuilding from a pak.ini

Private Sub OpenPrevious()
    Dim p As String, src As String
    p = PickFile(EngLibraryShare(), "Cartridge files (pak.ini)|pak.ini|Configuration settings (*.ini)|*.ini", "Select the pak.ini file of the cartridge to rebuild")
    If Len(p) = 0 Then Exit Sub
    If LCase$(Right$(p, 4)) <> ".ini" Then
        MsgBox "Select the pak.ini file in the cartridge folder.", vbExclamation, Caption
        Exit Sub
    End If
    If Not EngLoadPrevious(p) Then
        MsgBox "This file is not a valid pak.ini:" & vbCrLf & p, vbExclamation, Caption
        Exit Sub
    End If
    mPrevFile = p
    mPrevApplied = False
    mScanned = False
    src = EngPrev("source")
    mFilling = True
    txtSrc.Text = src
    mFilling = False
    ShowPage PG_FOLDER
    If EngFolderExists(src) Then
        lblPrev.Caption = "Loaded " & Squash(EngPrev("name") & " " & EngPrev("version")) & ". Scanning the game folder..."
        StartScan False
    Else
        lblPrev.Caption = "Loaded " & Squash(EngPrev("name") & " " & EngPrev("version")) & ". The game folder (" & src & ") was not found. Select its current location."
    End If
End Sub

Private Sub ApplyPrevious()
    Dim n As Long, i As Long
    mPrevApplied = True
    mName = EngPrev("name")
    mCartId = EngPrev("id")
    mVer = EngPrev("version")
    ' rebuilding: the next revision (a cartridge from before 0.9.2 counts as revision 1)
    n = Val(EngPrev("revision"))
    If n < 1 Then n = 1
    mRev = CStr(n + 1)
    mPlayers = EngPrev("players")
    mFolder = EngPrev("folder")
    mCartIdTouched = True
    mFolderTouched = True
    mOut = EngPrev("output")
    ' saved to the upload share before it was renamed: the new one
    If StrComp(mOut, EngOldShare(), vbTextCompare) = 0 Then mOut = EngServerShare()
    mReg = EngPrev("reg_source")
    mRegPath = EngPrev("reg_path")
    mRegAnyway = False
    mRegFull = ""
    If Len(mReg) > 0 Then CheckReg False
    mNotes = EngPrev("notes_source")
    mMore = ""
    mGCount = 0
    n = EngPrevGameCount()
    If n > MAX_GAMES Then n = MAX_GAMES
    mCands = n
    If n > 0 Then
        mPack = True
        mPackArgs = EngPrev("args")
        For i = 1 To n
            AddGameRow True, EngPrevGame(i, "name"), EngPrevGame(i, "players"), EngPrevGame(i, "exe"), EngPrevGame(i, "args"), _
                EngPrevGame(i, "start_in"), EngPrevGame(i, "icon_source"), EngPrevGame(i, "config_exe"), EngPrevGame(i, "config_args")
        Next
        ' games that had the cartridge's players show as "same as the cartridge"
        For i = 0 To mGCount - 1
            If mGPlayers(i) = mPlayers Then mGPlayers(i) = ""
        Next
    Else
        mPack = False
        mExe = EngPrev("exe")
        mArgs = EngPrev("args")
        mCfg = EngPrev("config_exe")
        mCfgArgs = EngPrev("config_args")
        mIcon = EngPrev("icon_source")
    End If
    If mGCount > 0 Then
        mSel = 0
    Else
        mSel = -1
    End If
    lblPrev.Caption = "Loaded " & Squash(mName & " " & mVer) & ". All settings have been loaded, and the revision has been increased to " & mRev & ". Change the version and any other settings, then click Next."
End Sub

' ================================================================ 2. details

Private Sub FillDetails()
    mFilling = True
    txtName.Text = mName
    txtId.Text = mCartId
    txtVer.Text = mVer
    txtRev.Text = mRev
    txtPlayers.Text = mPlayers
    txtFolder.Text = mFolder
    mFilling = False
End Sub

' The ID and folder follow the name, until they're changed by hand
Private Sub txtName_Change()
    If mFilling Then Exit Sub
    mName = Trim$(txtName.Text)
    mFilling = True
    If Not mCartIdTouched Then
        mCartId = EngMakeId(mName)
        txtId.Text = mCartId
    End If
    If Not mFolderTouched Then
        mFolder = EngMakeFolderName(mName)
        txtFolder.Text = mFolder
    End If
    mFilling = False
End Sub

Private Sub txtId_Change()
    If Not mFilling Then mCartIdTouched = True
End Sub

Private Sub txtFolder_Change()
    If Not mFilling Then mFolderTouched = True
End Sub

' ================================================================ 3. the games

Private Sub FillGames()
    Dim wasFilling As Boolean
    wasFilling = mFilling
    mFilling = True
    optSingle.Value = Not mPack
    optMulti.Value = mPack
    If mCands >= 2 And mPrevApplied Then
        lblMultiHint.Caption = "(" & mCands & " in previous build)"
    ElseIf mCands >= 2 Then
        lblMultiHint.Caption = "(" & mCands & " found)"
    Else
        lblMultiHint.Caption = ""
    End If
    fraSingle.Visible = Not mPack
    fraMulti.Visible = mPack
    FillPrograms cboExe, mExe, "(select)"
    txtArgs.Text = mArgs
    FillPrograms cboCfg, mCfg, "(none)"
    txtCfgArgs.Text = mCfgArgs
    txtCfgArgs.Enabled = (Len(mCfg) > 0)
    lblCfgArgs.Enabled = txtCfgArgs.Enabled
    txtIcon.Text = mIcon
    txtPackArgs.Text = mPackArgs
    FillGameList
    LoadGame
    mFilling = wasFilling
End Sub

Private Sub optSingle_Click()
    If mFilling Or Not mPack Then Exit Sub
    Collect PG_GAMES
    mPack = False
    FillGames
End Sub

Private Sub optMulti_Click()
    If mFilling Or mPack Then Exit Sub
    Collect PG_GAMES
    mPack = True
    ' start the multicart off with the one game so far
    If mGCount = 0 Then AddGameRow True, mName, "", mExe, mArgs, "", mIcon, mCfg, mCfgArgs
    If mSel < 0 Then mSel = 0
    FillGames
End Sub

' A program list: (select) or (none) first, then the programs in the game folder
Private Sub FillPrograms(cbo As ComboBox, ByVal current As String, ByVal noneText As String)
    Dim i As Long, parts() As String, want As Long
    cbo.Clear
    cbo.AddItem noneText
    For i = 0 To mExeCount - 1
        AddProgram cbo, mExes(i)
    Next
    If Len(mMore) > 0 Then
        parts = Split(mMore, "|")
        For i = 0 To UBound(parts)
            AddProgram cbo, parts(i)
        Next
    End If
    AddProgram cbo, current
    want = 0
    For i = 1 To cbo.ListCount - 1
        If StrComp(cbo.List(i), current, vbTextCompare) = 0 Then want = i
    Next
    cbo.ListIndex = want
End Sub

Private Sub AddProgram(cbo As ComboBox, ByVal p As String)
    Dim i As Long
    If Len(p) = 0 Then Exit Sub
    For i = 1 To cbo.ListCount - 1
        If StrComp(cbo.List(i), p, vbTextCompare) = 0 Then Exit Sub
    Next
    cbo.AddItem p
End Sub

Private Function ComboValue(cbo As ComboBox) As String
    If cbo.ListIndex > 0 Then ComboValue = cbo.List(cbo.ListIndex)
End Function

' A program deeper in the game folder than the list goes. True if one was picked.
Private Function OtherProgram(cbo As ComboBox, ByVal startIn As String) As Boolean
    Dim p As String, r As String, wasFilling As Boolean
    p = PickFile(startIn, "Programs (*.exe)|*.exe|All files|*.*", "Select the program")
    If Len(p) = 0 Then Exit Function
    r = EngRelPath(p)
    If r = p Or Not EngInGameFile(r) Then
        MsgBox "The program must be inside the game folder (" & mSrc & ").", vbExclamation, Caption
        Exit Function
    End If
    If InStr(1, "|" & mMore & "|", "|" & r & "|", vbTextCompare) = 0 Then
        If Len(mMore) > 0 Then mMore = mMore & "|"
        mMore = mMore & r
    End If
    wasFilling = mFilling
    mFilling = True
    FillPrograms cbo, r, cbo.List(0)
    mFilling = wasFilling
    OtherProgram = True
End Function

' An icon: inside the game it's kept as a path in it, anything else as it is
Private Function PickIcon(ByVal current As String, ByVal startIn As String) As String
    Dim p As String, r As String, full As String
    full = EngResolve(current)
    If Len(full) > 0 Then startIn = EngParentOf(full)
    p = PickFile(startIn, "Icons and images|*.ico;*.bmp;*.gif;*.jpg;*.jpeg|Programs and libraries|*.exe;*.dll|All files|*.*", "Select the icon")
    If Len(p) = 0 Then Exit Function
    If Not EngIsIconFile(p) Then
        MsgBox "Select an .ico, .bmp, .gif or .jpg file, or a program that contains an icon.", vbExclamation, Caption
        Exit Function
    End If
    r = EngRelPath(p)
    If r <> p And EngInGameFile(r) Then
        PickIcon = r
    Else
        PickIcon = p
    End If
End Function

' The folder a program is in, to start a picker there
Private Function ProgramDir(ByVal rel As String) As String
    If Len(rel) > 0 Then
        ProgramDir = EngParentOf(mSrc & "\" & rel)
    Else
        ProgramDir = mSrc
    End If
End Function

' The folder inside the game a multicart game's files are most likely in: its mod folder
Private Function GameDir(ByVal i As Long) As String
    Dim a As String, p As Long, m As String
    GameDir = mSrc
    If i < 0 Or i >= mGCount Then Exit Function
    If Len(mGStart(i)) > 0 And EngInGameFolder(mGStart(i)) Then
        GameDir = mSrc & "\" & mGStart(i)
        Exit Function
    End If
    a = " " & mGArgs(i) & " "
    p = InStr(1, a, " -game ", vbTextCompare)
    If p > 0 Then
        m = Trim$(Mid$(a, p + 7))
        If InStr(m, " ") > 0 Then m = Left$(m, InStr(m, " ") - 1)
        If EngInGameFolder(m) Then GameDir = mSrc & "\" & m
    End If
End Function

Private Sub cboExe_Click()
    If Not mFilling Then mExe = ComboValue(cboExe)
End Sub

Private Sub cboCfg_Click()
    If mFilling Then Exit Sub
    mCfg = ComboValue(cboCfg)
    txtCfgArgs.Enabled = (Len(mCfg) > 0)
    lblCfgArgs.Enabled = txtCfgArgs.Enabled
End Sub

Private Sub cmdExeBrowse_Click()
    If OtherProgram(cboExe, ProgramDir(ComboValue(cboExe))) Then mExe = ComboValue(cboExe)
End Sub

Private Sub cmdCfgBrowse_Click()
    If OtherProgram(cboCfg, ProgramDir(ComboValue(cboCfg))) Then
        mCfg = ComboValue(cboCfg)
        txtCfgArgs.Enabled = True
        lblCfgArgs.Enabled = True
    End If
End Sub

Private Sub cmdIconBrowse_Click()
    Dim p As String
    p = PickIcon(txtIcon.Text, mSrc)
    If Len(p) > 0 Then txtIcon.Text = p
End Sub

' ---- the games on a multicart

Private Sub AddGameRow(ByVal isOn As Boolean, ByVal nm As String, ByVal pl As String, ByVal ex As String, ByVal ar As String, _
                       ByVal st As String, ByVal ic As String, ByVal cf As String, ByVal cfa As String)
    If mGCount >= MAX_GAMES Then Exit Sub
    mGOn(mGCount) = isOn
    mGName(mGCount) = nm
    mGPlayers(mGCount) = pl
    mGExe(mGCount) = ex
    mGArgs(mGCount) = ar
    mGStart(mGCount) = st
    mGIcon(mGCount) = ic
    mGCfg(mGCount) = cf
    mGCfgArgs(mGCount) = cfa
    mGCount = mGCount + 1
End Sub

Private Sub FillGameList()
    Dim i As Long, wasFilling As Boolean
    wasFilling = mFilling
    mFilling = True
    lstGames.Clear
    For i = 0 To mGCount - 1
        lstGames.AddItem GameRowText(i)
        lstGames.Selected(lstGames.NewIndex) = mGOn(i)
    Next
    If mSel >= 0 And mSel < lstGames.ListCount Then lstGames.ListIndex = mSel
    cmdRemoveGame.Enabled = (mGCount > 0)
    mFilling = wasFilling
End Sub

' A game's row in the list: its name, and its players
Private Function GameRowText(ByVal i As Long) As String
    Dim s As String, pl As String
    s = mGName(i)
    If Len(s) = 0 Then s = "(unnamed)"
    pl = mGPlayers(i)
    If Len(pl) = 0 Then pl = mPlayers
    If Len(pl) > 0 Then s = s & "  (" & pl & ")"
    GameRowText = s
End Function

' The selected game, into the boxes on the right
Private Sub LoadGame()
    Dim ok As Boolean, wasFilling As Boolean
    wasFilling = mFilling
    mFilling = True
    ok = (mSel >= 0 And mSel < mGCount)
    txtGName.Enabled = ok
    txtGPlayers.Enabled = ok
    cboGExe.Enabled = ok
    cmdGExeBrowse.Enabled = ok
    txtGArgs.Enabled = ok
    txtGStart.Enabled = ok
    cmdGStartBrowse.Enabled = ok
    txtGIcon.Enabled = ok
    cmdGIconBrowse.Enabled = ok
    cboGCfg.Enabled = ok
    cmdGCfgBrowse.Enabled = ok
    If ok Then
        txtGName.Text = mGName(mSel)
        txtGPlayers.Text = mGPlayers(mSel)
        FillPrograms cboGExe, mGExe(mSel), "(select)"
        txtGArgs.Text = mGArgs(mSel)
        txtGStart.Text = mGStart(mSel)
        txtGIcon.Text = mGIcon(mSel)
        FillPrograms cboGCfg, mGCfg(mSel), "(none)"
        txtGCfgArgs.Text = mGCfgArgs(mSel)
        txtGCfgArgs.Enabled = (Len(mGCfg(mSel)) > 0)
    Else
        txtGName.Text = ""
        txtGPlayers.Text = ""
        cboGExe.Clear
        txtGArgs.Text = ""
        txtGStart.Text = ""
        txtGIcon.Text = ""
        cboGCfg.Clear
        txtGCfgArgs.Text = ""
        txtGCfgArgs.Enabled = False
    End If
    lblGCfgArgs.Enabled = txtGCfgArgs.Enabled
    ShowGameHints
    mFilling = wasFilling
End Sub

' The boxes on the right, into the selected game
Private Sub SaveGame()
    If mSel < 0 Or mSel >= mGCount Then Exit Sub
    mGName(mSel) = Trim$(txtGName.Text)
    mGPlayers(mSel) = Replace(Trim$(txtGPlayers.Text), " ", "")
    mGExe(mSel) = ComboValue(cboGExe)
    mGArgs(mSel) = Trim$(txtGArgs.Text)
    mGStart(mSel) = Trim$(txtGStart.Text)
    mGIcon(mSel) = Trim$(txtGIcon.Text)
    mGCfg(mSel) = ComboValue(cboGCfg)
    If Len(mGCfg(mSel)) > 0 Then
        mGCfgArgs(mSel) = Trim$(txtGCfgArgs.Text)
    Else
        mGCfgArgs(mSel) = ""
    End If
End Sub

' The whole command line the selected game starts with, and what a blank Players means
Private Sub ShowGameHints()
    If mSel >= 0 And mSel < mGCount Then
        lblGCmd.Caption = Squash(mGExe(mSel) & " " & mGArgs(mSel) & " " & Trim$(txtPackArgs.Text))
    Else
        lblGCmd.Caption = ""
    End If
    If Len(mPlayers) > 0 Then
        lblGPlayersHint.Caption = "Blank: same as the cartridge (" & mPlayers & ")"
    Else
        lblGPlayersHint.Caption = "Blank: same as the cartridge"
    End If
End Sub

' Something about the selected game changed: keep its row in the list up to date
Private Sub GameEdited()
    If mFilling Then Exit Sub
    SaveGame
    If mSel < 0 Or mSel >= lstGames.ListCount Then Exit Sub
    mFilling = True
    lstGames.List(mSel) = GameRowText(mSel)
    lstGames.Selected(mSel) = mGOn(mSel)
    mFilling = False
    ShowGameHints
End Sub

Private Sub lstGames_Click()
    If mFilling Then Exit Sub
    SaveGame
    mSel = lstGames.ListIndex
    LoadGame
End Sub

Private Sub lstGames_ItemCheck(Item As Integer)
    If mFilling Then Exit Sub
    If Item >= 0 And Item < mGCount Then mGOn(Item) = lstGames.Selected(Item)
End Sub

Private Sub txtGName_Change()
    GameEdited
End Sub

Private Sub txtGPlayers_Change()
    GameEdited
End Sub

Private Sub txtGArgs_Change()
    GameEdited
End Sub

Private Sub txtPackArgs_Change()
    If Not mFilling Then ShowGameHints
End Sub

Private Sub cboGExe_Click()
    GameEdited
End Sub

Private Sub cboGCfg_Click()
    If mFilling Then Exit Sub
    txtGCfgArgs.Enabled = (cboGCfg.ListIndex > 0)
    lblGCfgArgs.Enabled = txtGCfgArgs.Enabled
    GameEdited
End Sub

Private Sub cmdGExeBrowse_Click()
    If OtherProgram(cboGExe, GameDir(mSel)) Then GameEdited
End Sub

Private Sub cmdGCfgBrowse_Click()
    If OtherProgram(cboGCfg, GameDir(mSel)) Then
        txtGCfgArgs.Enabled = True
        lblGCfgArgs.Enabled = True
        GameEdited
    End If
End Sub

Private Sub cmdGStartBrowse_Click()
    Dim p As String, r As String
    p = EngPickFolder("Select the start folder:", mSrc)
    If Len(p) = 0 Then Exit Sub
    r = EngRelPath(p)
    If r = p Or StrComp(p, mSrc, vbTextCompare) = 0 Then
        MsgBox "Select a folder inside the game folder (" & mSrc & "), or leave the box blank.", vbExclamation, Caption
        Exit Sub
    End If
    txtGStart.Text = r
    GameEdited
End Sub

Private Sub cmdGIconBrowse_Click()
    Dim p As String
    p = PickIcon(txtGIcon.Text, GameDir(mSel))
    If Len(p) = 0 Then Exit Sub
    txtGIcon.Text = p
    GameEdited
End Sub

Private Sub cmdAddGame_Click()
    SaveGame
    If mGCount >= MAX_GAMES Then
        MsgBox "A multicart can contain up to " & MAX_GAMES & " games.", vbExclamation, Caption
        Exit Sub
    End If
    If mExeCount > 0 Then
        AddGameRow True, "", "", EngGuessExe(mName, mCartId), "", "", "", "", ""
    Else
        AddGameRow True, "", "", "", "", "", "", "", ""
    End If
    mSel = mGCount - 1
    FillGameList
    LoadGame
    On Error Resume Next
    txtGName.SetFocus
End Sub

Private Sub cmdRemoveGame_Click()
    Dim i As Long, nm As String
    SaveGame
    If mSel < 0 Or mSel >= mGCount Then Exit Sub
    nm = mGName(mSel)
    If Len(nm) = 0 Then nm = "this game"
    If MsgBox("Remove " & nm & " from the list?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) <> vbYes Then Exit Sub
    For i = mSel To mGCount - 2
        mGOn(i) = mGOn(i + 1)
        mGName(i) = mGName(i + 1)
        mGPlayers(i) = mGPlayers(i + 1)
        mGExe(i) = mGExe(i + 1)
        mGArgs(i) = mGArgs(i + 1)
        mGStart(i) = mGStart(i + 1)
        mGIcon(i) = mGIcon(i + 1)
        mGCfg(i) = mGCfg(i + 1)
        mGCfgArgs(i) = mGCfgArgs(i + 1)
    Next
    mGCount = mGCount - 1
    If mSel >= mGCount Then mSel = mGCount - 1
    FillGameList
    LoadGame
End Sub

' ================================================================ 4. registry file and setup notes

Private Sub FillExtras()
    Dim i As Long
    mFilling = True
    cboReg.Clear
    For i = 0 To mRegCount - 1
        cboReg.AddItem mRegsFound(i)
    Next
    cboReg.Text = mReg
    txtRegPath.Text = mRegPath
    If mRegAnyway Then
        chkRegAnyway.Value = vbChecked
    Else
        chkRegAnyway.Value = vbUnchecked
    End If
    txtNotes.Text = mNotes
    mFilling = False
    RenderReg
End Sub

' A different registry file: read it now, and suggest its install path
Private Sub RegChanged()
    mReg = Trim$(cboReg.Text)
    mRegAnyway = False
    mRegPath = ""
    mRegFull = ""
    If Len(mReg) > 0 Then CheckReg True
    mFilling = True
    txtRegPath.Text = mRegPath
    chkRegAnyway.Value = vbUnchecked
    mFilling = False
    RenderReg
End Sub

Private Sub CheckReg(ByVal guessPath As Boolean)
    mRegFull = EngCheckReg(mReg)
    If Len(mRegFull) = 0 Then Exit Sub
    mRegKeys = EngRegInfo("keys")
    mRegValues = EngRegInfo("values")
    mRegNotes = EngRegInfo("notes")
    mRegTrouble = EngRegInfo("trouble")
    If guessPath And Len(mRegPath) = 0 Then mRegPath = EngRegInfo("guess")
End Sub

Private Sub RenderReg()
    Dim s As String, i As Long, hasReg As Boolean
    If Len(mReg) = 0 Then
        s = "Optional."
        If mRegCount > 0 Then
            s = s & " Found in the game folder: "
            For i = 0 To mRegCount - 1
                If i > 0 Then s = s & ", "
                s = s & mRegsFound(i)
            Next
            s = s & "."
        End If
        lblRegInfo.ForeColor = C_GREY
        lblRegInfo.Caption = s
        chkRegAnyway.Visible = False
    ElseIf Len(mRegFull) = 0 Then
        lblRegInfo.ForeColor = C_RED
        lblRegInfo.Caption = "Registry file not found."
        chkRegAnyway.Visible = False
    Else
        s = mRegKeys & " key" & Plural(mRegKeys) & ", " & mRegValues & " value" & Plural(mRegValues) & "."
        If Len(mRegNotes) > 0 Then s = s & " " & Replace(mRegNotes, vbLf, " ")
        If Len(mRegTrouble) > 0 Then
            s = s & " Errors: " & Replace(mRegTrouble, vbLf, " ")
            lblRegInfo.ForeColor = C_RED
        Else
            lblRegInfo.ForeColor = vbButtonText
        End If
        lblRegInfo.Caption = s
        chkRegAnyway.Visible = (Len(mRegTrouble) > 0)
    End If
    hasReg = (Len(mRegFull) > 0)
    lblRegPath.Enabled = hasReg
    txtRegPath.Enabled = hasReg
    If hasReg And Len(mRegPath) > 0 And Not EngRegMentions(mRegPath) Then
        lblRegPathHint.ForeColor = C_RED
        lblRegPathHint.Caption = "This path does not appear in the registry file. No values will be replaced."
    Else
        lblRegPathHint.ForeColor = C_GREY
        lblRegPathHint.Caption = REG_PATH_HINT
    End If
End Sub

Private Sub cboReg_Click()
    If mFilling Then Exit Sub
    If cboReg.ListIndex >= 0 Then
        mFilling = True
        cboReg.Text = cboReg.List(cboReg.ListIndex)
        mFilling = False
    End If
    RegChanged
End Sub

Private Sub cboReg_LostFocus()
    If StrComp(Trim$(cboReg.Text), mReg, vbTextCompare) <> 0 Then RegChanged
End Sub

Private Sub chkRegAnyway_Click()
    If Not mFilling Then mRegAnyway = (chkRegAnyway.Value = vbChecked)
End Sub

Private Sub txtRegPath_Change()
    If mFilling Then Exit Sub
    mRegPath = Trim$(txtRegPath.Text)
    RenderReg
End Sub

Private Sub cmdRegBrowse_Click()
    Dim p As String, startIn As String
    startIn = mSrc
    If Len(EngResolve(mReg)) > 0 Then startIn = EngParentOf(EngResolve(mReg))
    p = PickFile(startIn, "Registry files (*.reg)|*.reg|All files|*.*", "Select the registry file")
    If Len(p) = 0 Then Exit Sub
    mFilling = True
    cboReg.Text = p
    mFilling = False
    RegChanged
End Sub

Private Sub cmdNotesBrowse_Click()
    Dim p As String, startIn As String
    startIn = mSrc
    If Len(EngResolve(txtNotes.Text)) > 0 Then startIn = EngParentOf(EngResolve(txtNotes.Text))
    p = PickFile(startIn, "Text files (*.txt)|*.txt|All files|*.*", "Select the setup notes")
    If Len(p) > 0 Then txtNotes.Text = p
End Sub

Private Sub cmdNotesView_Click()
    Dim p As String
    p = EngResolve(Trim$(txtNotes.Text))
    If Len(p) > 0 Then
        EngOpenFile p
    Else
        MsgBox "Select a text file first.", vbExclamation, Caption
    End If
End Sub

' ================================================================ checking the answers

' Shows the first problem on a page and goes to its box. True if there are none.
Private Function PageOK(ByVal n As Integer) As Boolean
    Dim s As String
    s = PageProblems(n, True)
    If Len(s) = 0 Then
        PageOK = True
        Exit Function
    End If
    MsgBox s, vbExclamation, Caption
    On Error Resume Next
    If Not mFocus Is Nothing Then mFocus.SetFocus
End Function

' What's wrong on a page, one per line. firstOnly: just the first, and mFocus set to its box.
Private Function PageProblems(ByVal n As Integer, ByVal firstOnly As Boolean) As String
    Dim out As String, i As Long, s As String, t As String, anyOn As Boolean
    Set mFocus = Nothing
    Select Case n
    Case PG_DETAILS
        If Len(mName) = 0 Then AddProblem out, "Enter a name for the game.", txtName
        If Not EngIsId(mCartId) Then AddProblem out, "The ID must be 2 to 24 lowercase letters, numbers or - (for example, ut99).", txtId
        If Not EngIsPlayers(mPlayers) Then AddProblem out, "Enter the number of players: a number (16) or a range (2-16).", txtPlayers
        If Not RevOK(mRev) Then AddProblem out, "The revision must be a whole number from 1 to 999.", txtRev
        If Len(mFolder) = 0 Or EngMakeFolderName(mFolder) <> mFolder Then AddProblem out, "The install folder name can contain letters, numbers, -, _ and . only.", txtFolder
    Case PG_GAMES
        If Not mPack Then
            If Len(mExe) = 0 Then
                AddProblem out, "Select the program that starts the game.", cboExe
            ElseIf Not EngInGameFile(mExe) Then
                AddProblem out, mExe & " was not found in the game folder.", cboExe
            End If
            If Len(mCfg) > 0 And Not EngInGameFile(mCfg) Then AddProblem out, mCfg & " was not found in the game folder.", cboCfg
            s = IconProblem(mIcon)
            If Len(s) > 0 Then AddProblem out, s, txtIcon
        Else
            For i = 0 To mGCount - 1
                If mGOn(i) Then
                    anyOn = True
                    s = GameProblem(i)
                    If Len(s) > 0 Then
                        t = mGName(i)
                        If Len(t) = 0 Then t = "Unnamed game"
                        If firstOnly And Len(out) = 0 Then
                            ' show that game, so its box can be corrected
                            mSel = i
                            FillGameList
                            LoadGame
                        End If
                        AddProblem out, t & ": " & s, txtGName
                    End If
                End If
            Next
            If Not anyOn Then AddProblem out, "Select at least one game for the multicart.", lstGames
        End If
    Case PG_EXTRAS
        If Len(mReg) > 0 Then
            If Len(EngResolve(mReg)) = 0 Then
                AddProblem out, "Registry file not found: " & mReg, cboReg
            ElseIf Len(mRegTrouble) > 0 And Not mRegAnyway Then
                AddProblem out, "The registry file has errors. Select Use this file anyway, or select another file.", chkRegAnyway
            End If
        End If
        If Len(mNotes) > 0 And Len(EngResolve(mNotes)) = 0 Then AddProblem out, "Setup notes file not found: " & mNotes, txtNotes
    Case PG_READY
        If Len(mOut) = 0 Then AddProblem out, "Select a folder to save the cartridge in.", txtOut
    End Select
    If firstOnly And InStr(out, vbCrLf) > 0 Then out = Left$(out, InStr(out, vbCrLf) - 1)
    PageProblems = out
End Function

Private Sub AddProblem(out As String, ByVal s As String, ByVal box As Object)
    If Len(out) = 0 Then Set mFocus = box
    If Len(out) > 0 Then out = out & vbCrLf
    out = out & s
End Sub

Private Function GameProblem(ByVal i As Long) As String
    Dim s As String
    If Len(mGName(i)) = 0 Then
        s = "Enter a name."
    ElseIf Len(mGPlayers(i)) > 0 And Not EngIsPlayers(mGPlayers(i)) Then
        s = "Players must be a number (16) or a range (2-16)."
    ElseIf Len(mGExe(i)) = 0 Then
        s = "Select the program that starts the game."
    ElseIf Not EngInGameFile(mGExe(i)) Then
        s = mGExe(i) & " was not found in the game folder."
    ElseIf Len(mGStart(i)) > 0 And Not EngInGameFolder(mGStart(i)) Then
        s = "The start folder " & mGStart(i) & " was not found in the game folder."
    ElseIf Len(mGCfg(i)) > 0 And Not EngInGameFile(mGCfg(i)) Then
        s = mGCfg(i) & " was not found in the game folder."
    Else
        s = IconProblem(mGIcon(i))
    End If
    GameProblem = s
End Function

Private Function IconProblem(ByVal ic As String) As String
    If Len(ic) = 0 Then Exit Function
    If Len(EngResolve(ic)) = 0 Then
        IconProblem = "Icon file not found: " & ic & "."
    ElseIf Not EngIsIconFile(ic) Then
        IconProblem = "The icon must be an .ico, .bmp, .gif or .jpg file, or a program."
    End If
End Function

' ================================================================ 5. ready to make

Private Sub FillReady()
    mFilling = True
    txtOut.Text = mOut
    mFilling = False
    If Len(mOut) = 0 And Not mOutChecked Then
        ' looking for the Cartridge Server can take a few seconds when it isn't there: after the page is up
        mOutChecked = True
        lblOutHint.Caption = "Searching for the Cartridge Server..."
        mJob = "server"
        tmrWork.Interval = 50
        tmrWork.Enabled = True
    End If
    RenderReady
End Sub

Private Sub ServerTick()
    tmrWork.Enabled = False
    Screen.MousePointer = vbHourglass
    mOut = EngDefaultOutput()
    Screen.MousePointer = vbDefault
    mJob = ""
    mFilling = True
    txtOut.Text = mOut
    mFilling = False
    RenderReady
    UpdateButtons
End Sub

Private Sub txtOut_Change()
    If mFilling Then Exit Sub
    mOut = Trim$(txtOut.Text)
    RenderReady
    UpdateButtons
End Sub

Private Sub cmdOutBrowse_Click()
    Dim p As String
    p = EngPickFolder("Select the folder to save the cartridge in:", "")
    If Len(p) > 0 Then txtOut.Text = p
End Sub

' The summary, and anything to sort out first
Private Sub RenderReady()
    Dim probs As String, k As Integer, s As String
    If mJob = "server" Then
        mReadyOK = False
    ElseIf Len(mOut) = 0 Then
        lblOutHint.Caption = "Select a folder to save the cartridge in."
    ElseIf EngOnServer(mOut) Then
        lblOutHint.Caption = "Cartridge Server upload share. The host approves the cartridge before it appears in the Game Library."
    Else
        lblOutHint.Caption = "A folder on this PC. Copy the cartridge to " & EngServerShare() & " after it is made."
    End If
    PushToEngine
    DrawSummary
    ' every page's problems, then this one's
    For k = PG_DETAILS To PG_READY
        s = PageProblems(k, False)
        If Len(s) > 0 Then
            If Len(probs) > 0 Then probs = probs & vbCrLf
            probs = probs & s
        End If
    Next
    If Len(probs) > 0 Then
        lblProblems.ForeColor = C_RED
        lblProblems.Caption = "Correct the following before making the cartridge (click Back):" & vbCrLf & probs
    Else
        lblProblems.ForeColor = vbButtonText
        lblProblems.Caption = Warnings()
    End If
    mReadyOK = (Len(probs) = 0) And mJob <> "server"
End Sub

' Things to know before making it
Private Function Warnings() As String
    Dim w As String, lp() As String, freeBytes As Double, pd As String, nm As String
    lp = Split(EngLongPaths(mFolder), vbLf)
    If Val(lp(0)) > 0 Then
        AddLine w, lp(0) & " file" & Plural(Val(lp(0))) & " will exceed the Windows path length limit when installed, for example " & lp(1) & ". Use a shorter folder name."
    End If
    If Len(mOut) > 0 And EngFolderExists(mOut) Then
        freeBytes = EngFreeSpace(mOut)
        If freeBytes >= 0 And freeBytes < EngScanInfo("bytes") + 10485760 Then
            AddLine w, "Insufficient disk space: " & EngSizeText(freeBytes) & " free, approximately " & mSizeText & " required."
        End If
    End If
    pd = EngPakDir()
    nm = Mid$(pd, InStrRev(pd, "\") + 1)
    If EngOnServer(mOut) And Len(mCartId) > 0 Then
        If EngLibraryHas(nm) Then AddLine w, "An approved cartridge named " & nm & " is already in the Game Library. It is replaced when this cartridge is approved."
    End If
    If Len(mOut) > 0 And Len(mCartId) > 0 Then
        If EngPakExists() Then AddLine w, "A cartridge already exists in " & pd & ". It will be replaced."
    End If
    Warnings = w
End Function

Private Sub AddLine(s As String, ByVal more As String)
    If Len(s) > 0 Then s = s & vbCrLf
    s = s & more
End Sub

' A revision: a whole number from 1 to 999
Private Function RevOK(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) < 1 Or Len(s) > 3 Then Exit Function
    For i = 1 To Len(s)
        If InStr("0123456789", Mid$(s, i, 1)) = 0 Then Exit Function
    Next
    RevOK = (Val(s) >= 1)
End Function

' Every answer, into the engine
Private Sub PushToEngine()
    Dim i As Long
    EngSet "name", mName
    EngSet "id", mCartId
    EngSet "version", mVer
    EngSet "revision", CStr(Val(mRev))
    EngSet "players", mPlayers
    EngSet "folder", mFolder
    If mPack Then
        EngSet "pack", "1"
        EngSet "packargs", mPackArgs
        EngSet "exe", ""
        EngSet "args", ""
        EngSet "cfg", ""
        EngSet "cfgargs", ""
        EngSet "icon", ""
    Else
        EngSet "pack", "0"
        EngSet "packargs", ""
        EngSet "exe", mExe
        EngSet "args", mArgs
        EngSet "cfg", mCfg
        EngSet "cfgargs", mCfgArgs
        EngSet "icon", mIcon
    End If
    If Len(mReg) > 0 Then
        EngSet "reg", EngResolve(mReg)
        EngSet "regpath", mRegPath
    Else
        EngSet "reg", ""
        EngSet "regpath", ""
    End If
    If Len(mNotes) > 0 Then
        EngSet "notes", EngResolve(mNotes)
    Else
        EngSet "notes", ""
    End If
    EngSet "output", mOut
    EngClearGames
    If mPack Then
        For i = 0 To mGCount - 1
            If mGOn(i) Then EngAddGame mGName(i), mGPlayers(i), mGExe(i), mGArgs(i), mGStart(i), mGIcon(i), mGCfg(i), mGCfgArgs(i)
        Next
    End If
End Sub

Private Sub DrawSummary()
    Dim y As Long, i As Long, s As String, n As Long
    picSum.Cls
    y = 6
    SumRow y, "Game", Squash(mName & " " & VerText(mVer, mRev)) & " (" & mCartId & ", " & mPlayers & " players)", True
    If mPack Then
        For i = 0 To mGCount - 1
            If mGOn(i) Then
                If Len(s) > 0 Then s = s & ", "
                s = s & mGName(i)
                n = n + 1
            End If
        Next
        SumRow y, "Multicart", n & " game" & Plural(n) & ": " & s, False
        If Len(mPackArgs) > 0 Then SumRow y, "All games", mPackArgs, False
    Else
        s = Squash(mExe & " " & mArgs)
        If Len(mCfg) > 0 Then s = s & " (configuration: " & Squash(mCfg & " " & mCfgArgs) & ")"
        SumRow y, "Command line", s, False
        If Len(mIcon) > 0 Then
            SumRow y, "Icon", mIcon, False
        Else
            SumRow y, "Icon", "from " & mExe, False
        End If
    End If
    If Len(mReg) > 0 Then
        s = mReg
        If Len(mRegPath) > 0 Then s = s & " (install path: " & mRegPath & ")"
        SumRow y, "Registry file", s, False
    Else
        SumRow y, "Registry file", "none", False
    End If
    If Len(mNotes) > 0 Then
        SumRow y, "Setup notes", mNotes, False
    Else
        SumRow y, "Setup notes", "none", False
    End If
    SumRow y, "Install folder", EngGamesRoot() & mFolder, False
    SumRow y, "Size", NumText(mFiles) & " files, " & mSizeText & ", in " & mPieces & " piece" & Plural(mPieces), False
    If Len(mOut) > 0 And Len(mCartId) > 0 Then SumRow y, "Cartridge", EngPakDir(), False
End Sub

Private Sub SumRow(y As Long, ByVal title As String, ByVal v As String, ByVal bold As Boolean)
    Dim w As Long
    w = picSum.ScaleWidth - 108
    TextAt picSum, 8, y, title, C_SUB
    If bold Then picSum.FontBold = True
    v = FitText(picSum, v, w)
    picSum.FontBold = False
    TextAt picSum, 100, y, v, vbBlack, bold
    y = y + 18
End Sub

' ================================================================ 6. making it

Private Sub Make()
    Dim r As String
    RenderReady
    If Not mReadyOK Then
        UpdateButtons
        Exit Sub
    End If
    PushToEngine
    If Not EngMakeFolder(mOut) Then
        MsgBox "Unable to create or open the folder " & mOut & ".", vbExclamation, Caption
        Exit Sub
    End If
    If EngPakExists() Then
        If MsgBox("A cartridge already exists in" & vbCrLf & EngPakDir() & vbCrLf & vbCrLf & "Replace it?", vbQuestion + vbYesNo + vbDefaultButton2, Caption) <> vbYes Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    r = EngBuildStart(True)
    Screen.MousePointer = vbDefault
    If Len(r) > 0 Then
        mResult = "fail|" & r
        mBuilt = False
        ShowPage PG_DONE
        Exit Sub
    End If
    mBuilding = True
    mStopping = False
    mBuildStart = GetTickCount()
    mPct = 0
    lblMaking.Caption = "Making " & mName & "..."
    lblMakingSub.Caption = "Starting 7-Zip..."
    ShowPage PG_MAKING
    DrawProgress
    mJob = "build"
    tmrWork.Interval = 1000
    tmrWork.Enabled = True
End Sub

Private Sub BuildTick()
    Dim r As String, parts() As String, secs As Long, detail As String
    r = EngBuildPoll()
    parts = Split(r & "|", "|")
    secs = (GetTickCount() - mBuildStart) \ 1000
    If secs < 0 Then secs = 0
    If parts(0) = "packing" Then
        mPct = Val(parts(1))
        DrawProgress
        If mStopping Then
            lblMaking.Caption = "Stopping..."
        Else
            lblMaking.Caption = "Making " & mName & ": " & mPct & "%"
        End If
        detail = "Elapsed: " & MinSec(secs)
        If mPct >= 3 Then detail = detail & ", remaining: about " & MinSec((secs * (100 - mPct)) \ mPct)
        If Val(parts(2)) > 0 Then detail = detail & ". File " & NumText(Val(parts(2))) & " of " & NumText(Val(parts(3))) & "."
        lblMakingSub.Caption = detail
        Exit Sub
    End If
    tmrWork.Enabled = False
    mJob = ""
    mBuilding = False
    mStopping = False
    mResult = r
    mBuilt = (parts(0) = "done")
    mBuildSecs = secs
    ShowPage PG_DONE
End Sub

Private Sub StopBuild()
    If Not mBuilding Or mStopping Then Exit Sub
    If MsgBox("Stop making the cartridge? The incomplete cartridge will be removed.", vbQuestion + vbYesNo + vbDefaultButton2, Caption) <> vbYes Then Exit Sub
    mStopping = True
    UpdateButtons
    lblMaking.Caption = "Stopping..."
    EngBuildStop
End Sub

Private Sub DrawProgress()
    picProg.Cls
    ProgressBar picProg, 0, 0, picProg.ScaleWidth, picProg.ScaleHeight, mPct
End Sub

' ================================================================ 7. done

Private Sub FillDone()
    Dim s As String, note As String
    cmdApprove.Visible = False
    cmdOpenFolder.Visible = False
    cmdAnother.Visible = mBuilt
    If mBuilt Then
        lblDTitle.Caption = "Completing the Make Cartridge Wizard"
        s = Squash(mName & " " & VerText(mVer, mRev)) & " was made (" & MinSec(mBuildSecs) & "). "
        If EngOnServer(mOut) Then
            s = s & "It was saved to the Cartridge Server, and is moved to the approval queue within a few seconds. "
            If frmMain.HostMode() Then
                s = s & "Click Approve Cartridge to approve it. It then appears in the Game Library on every PC."
                cmdApprove.Visible = True
            Else
                s = s & "The host must approve it before it appears in the Game Library."
            End If
        Else
            s = s & "It was saved in " & EngPakDir() & ". Copy the cartridge folder to " & EngServerShare() & ". The host must approve it before it appears in the Game Library."
            cmdOpenFolder.Visible = True
        End If
        note = Mid$(mResult, 6)
        If Len(note) > 0 Then s = s & vbCrLf & vbCrLf & "Note: " & Replace(note, vbLf, vbCrLf)
        lblDNext.Caption = "To close this wizard, click Finish."
    ElseIf mResult = "stopped" Then
        lblDTitle.Caption = "The Cartridge Was Not Made"
        s = "Making the cartridge was cancelled. The incomplete cartridge was removed."
        lblDNext.Caption = "To try again, click Back. To close this wizard, click Finish."
    Else
        lblDTitle.Caption = "The Cartridge Could Not Be Made"
        s = Replace(Mid$(mResult, 6), vbLf, vbCrLf)
        lblDNext.Caption = "To change the settings and try again, click Back."
    End If
    lblDText.Caption = s
End Sub

' The host: straight to the Cartridge Server's administration page
Private Sub cmdApprove_Click()
    EngOpenUrl EngAdminUrl()
End Sub

Private Sub cmdOpenFolder_Click()
    EngOpenFolder EngPakDir()
End Sub

Private Sub cmdAnother_Click()
    StartOver
End Sub

' ================================================================ small helpers

' The Windows Open box, starting in a folder of our choosing. "" if they cancelled.
' filt is like "Programs (*.exe)|*.exe|All files|*.*"
Private Function PickFile(ByVal startIn As String, ByVal filt As String, ByVal title As String) As String
    Dim ofn As OPENFILENAME, n As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.lpstrFilter = Replace(filt, "|", vbNullChar) & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = String$(520, vbNullChar)
    ofn.nMaxFile = 520
    If Len(startIn) > 0 Then
        If EngFolderExists(startIn) Then ofn.lpstrInitialDir = startIn
    End If
    ofn.lpstrTitle = title
    ofn.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR
    If GetOpenFileName(ofn) = 0 Then Exit Function
    n = InStr(ofn.lpstrFile, vbNullChar)
    If n > 0 Then
        PickFile = Left$(ofn.lpstrFile, n - 1)
    Else
        PickFile = ofn.lpstrFile
    End If
End Function

Private Function NumText(ByVal n As Double) As String
    NumText = Format$(n, "#,##0")
End Function

Private Function Plural(ByVal n As Double) As String
    If n <> 1 Then Plural = "s"
End Function

Private Function MinSec(ByVal secs As Long) As String
    MinSec = (secs \ 60) & ":" & Format$(secs Mod 60, "00")
End Function

' One space between words
Private Function Squash(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squash = s
End Function
