VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Location"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMain.frx":0000
   ScaleHeight     =   588
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrFirstdraw 
      Interval        =   1
      Left            =   8160
      Top             =   6600
   End
   Begin VB.Frame fmequestions 
      BackColor       =   &H006B0500&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Answer1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2400
         Width           =   4440
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Answer2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3000
         Width           =   4440
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Answer3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3600
         Width           =   4440
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Answer4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4200
         Width           =   4440
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Correct Answer"
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   12
         Top             =   7755
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Question"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   5
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   9480
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Question Number"
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   10
         Top             =   8160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdBegin 
         Caption         =   "Begin..."
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtUserAnswer 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4800
         Width           =   4440
      End
      Begin VB.CommandButton ConfirmAnswer 
         Caption         =   "Confirm..."
         Height          =   300
         Left            =   8520
         TabIndex        =   7
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAn4 
         Caption         =   "Choose..."
         Height          =   300
         Left            =   8520
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAn3 
         Caption         =   "Choose..."
         Height          =   300
         Left            =   8520
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAn2 
         Caption         =   "Choose..."
         Height          =   300
         Left            =   8520
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAn1 
         Caption         =   "Choose..."
         Height          =   300
         Left            =   8520
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer tmrupdate 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   3120
         Top             =   5880
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   29
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   28
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer 3:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   27
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer 4:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   26
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Correct Answer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   25
         Top             =   7755
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Question:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Question Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   8160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image imgPortrait 
         BorderStyle     =   1  'Fixed Single
         Height          =   2820
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2310
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "You Say:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   22
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Question: 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label lblCorrect 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Correct:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label lblWrong 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wrong:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label lblresult 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click ""Begin"" to start the Knowledge Quest..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   5280
         Width           =   8775
      End
      Begin VB.Label lblCurrentLevel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Level:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Image Banner 
         Height          =   3000
         Left            =   0
         Picture         =   "FrmMain.frx":A915
         Top             =   0
         Width           =   30000
      End
   End
   Begin VB.TextBox txtConverse 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmMain.frx":12F8D9
      Top             =   4920
      Width           =   12420
   End
   Begin VB.Line loadbar4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   32
      X2              =   33
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Image left2 
      Height          =   480
      Left            =   7200
      Picture         =   "FrmMain.frx":12F916
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image right2 
      Height          =   480
      Left            =   7680
      Picture         =   "FrmMain.frx":13055A
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgupleft 
      Height          =   480
      Left            =   7200
      Picture         =   "FrmMain.frx":13119E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgdownleft 
      Height          =   480
      Left            =   7200
      Picture         =   "FrmMain.frx":131DE2
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEl 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMain.frx":132A26
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgm 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMain.frx":13366A
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgtorch 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMain.frx":1342AE
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgwindow 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMain.frx":134EF2
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgdoors 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMain.frx":135B36
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgwall5 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMain.frx":13677A
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWall4 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMain.frx":1373BE
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgwall3 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMain.frx":138002
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgleftedgewall 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMain.frx":138C46
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgwall 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMain.frx":13988A
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStoneFloor 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMain.frx":13A4CE
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallRight 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMain.frx":13B112
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallLeft 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMain.frx":13BD56
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWallback 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMain.frx":13C99A
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFrontWall 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMain.frx":13D5DE
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image down 
      Height          =   480
      Left            =   7680
      Picture         =   "FrmMain.frx":13E222
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image right 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMain.frx":13EE66
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image left1 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMain.frx":13FAAA
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image up 
      Height          =   480
      Left            =   7680
      Picture         =   "FrmMain.frx":1406EE
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image onbridge1up 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMain.frx":141332
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgonbridgedown1 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMain.frx":141F76
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image mgonbridge2down 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMain.frx":142BBA
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgonbridge2up 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMain.frx":1437FE
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgonbridge3down 
      Height          =   480
      Left            =   960
      Picture         =   "FrmMain.frx":144442
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgonbridge3up 
      Height          =   480
      Left            =   480
      Picture         =   "FrmMain.frx":145086
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbrifgenorth 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMain.frx":145CCA
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image imgTreestop 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMain.frx":14690E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBlack 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMain.frx":147552
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBridge 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMain.frx":148196
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIlt 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMain.frx":148DDA
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIrb 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMain.frx":149A1E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIrt 
      Height          =   480
      Left            =   2880
      Picture         =   "FrmMain.frx":14A662
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IMgVillage 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMain.frx":14B2A6
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSign 
      Height          =   480
      Left            =   0
      Picture         =   "FrmMain.frx":14BEEA
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgSkull 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMain.frx":14CB2E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTl 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMain.frx":14D772
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTori 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMain.frx":14E3B6
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTr 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMain.frx":14EFFA
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgTT 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMain.frx":14FC3E
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgUpG 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMain.frx":150882
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMount3 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMain.frx":1514C6
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMount2 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMain.frx":15210A
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGuy5 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmMain.frx":152D4E
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMount1 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMain.frx":153992
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgLG 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMain.frx":1545D6
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgL 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMain.frx":15521A
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgR 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMain.frx":155E5E
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMount4 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMain.frx":156AA2
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgRG 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMain.frx":1576E6
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGuy3 
      Height          =   480
      Left            =   4800
      Picture         =   "FrmMain.frx":15832A
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgDownG 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMain.frx":158F6E
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGrass 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMain.frx":159BB2
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgForest 
      Height          =   480
      Left            =   6240
      Picture         =   "FrmMain.frx":15A7F6
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGuy1 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMain.frx":15B43A
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMountain 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMain.frx":15C07E
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgIlb 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMain.frx":15CCC2
      Top             =   7080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGuy4 
      Height          =   480
      Left            =   5280
      Picture         =   "FrmMain.frx":15D906
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGuy2 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmMain.frx":15E54A
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBush 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmMain.frx":15F18E
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBr 
      Height          =   480
      Left            =   2400
      Picture         =   "FrmMain.frx":15FDD2
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBl 
      Height          =   480
      Left            =   1440
      Picture         =   "FrmMain.frx":160A16
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgBB 
      Height          =   480
      Left            =   1920
      Picture         =   "FrmMain.frx":16165A
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   32
      X2              =   664
      Y1              =   160
      Y2              =   160
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter1 As Integer
Dim Look As String

Private Sub cmdBegin_Click()
    updatetext
    cmdBegin.Visible = False
End Sub

Sub updatetext()

counter2 = 0

'Although the questions are displayed on this form they are not loaded on this form, they
'...are actually loaded on frmQuestions. the frmquestions does all the database linking and loading
'...This sub sinmply synchronises the correct text boxes and labels with the correct data
'...Rather than call the text boxes txtquestion1 etc, i put them all in an array so tghey can be
'...synchronised using a 3 line loop.

Do
    txtFields(counter2).Text = frmQuestions.txtFields(counter2)
    counter2 = counter2 + 1
Loop Until counter2 = 6

'make a few select butons visible beause the user has clicked begin

    cmdAn1.Visible = True
    cmdAn2.Visible = True
    cmdAn3.Visible = True
    cmdAn4.Visible = True
    ConfirmAnswer.Visible = True
'The form has a few labels telling the user about their progress, the following few lines fill them in
    loopcounter = loopcounter + 1
    currentquestion = currentquestion + 1
    lblProgress.Caption = "Question: " & currentquestion & " \ 10"
    lblCorrect.Caption = "Correct: " & correcttemp & ""
    lblWrong.Caption = "Wrong: " & wrongtemp & ""
    lblCurrentLevel.Caption = "Current Level: " & currentlevel & ""
    txtUserAnswer.Text = ""
'when the user does something, a label displays a comment like correct or incorrect, the aswer was actually etc...
'...this label will go blank however after about 2 seconds, so i used a simple timer to clear the feild
    tmrupdate.Enabled = True
    
'This code is responsible for ending the set of ten questions, and deciding if they should fight or not
'...it also monitors current level, if they have beaten the last question master, then it calls the game over form
'...It decides if a battle is nessasery by comparing the amount of correct answers, to avoid a fight one must get more right than their current level.
'...for example, if on level 2, you need to get three questions correct.

    If currentquestion >= 11 Then
        If correcttemp > currentlevel Then
            EnterMessageText = "You needed more than " & currentlevel & " correct answer(s), you got " & correcttemp & ". Well done, A prize you shall receive!"
'OKMessage is another one of my custom msgbox's
            OKMessage
            If currentlevel = 2 Then
                gameoverreason = "donequestions"
                gameover
                Exit Sub
            Else
                issueprize
            End If
        Else
            EnterMessageText = "You needed more than " & currentlevel & " correct answer(s), you only got " & correcttemp & ". You shall pay for you're impotence!"
            OKMessage
            If currentlevel = 2 Then
            EnterMessageText = "You need to defeat this question master in a quiz to complete the game!"
            OKMessage
            Else
                issueprize
                frmWorldMap.shpBlankout.Width = frmWorldMap.ScaleWidth
                frmWorldMap.shpBlankout.Height = frmWorldMap.ScaleHeight
                frmWorldMap.shpBlankout.Visible = True
                Unload frmbATTLE
                frmbATTLE.Show
            End If
        End If
    End If
    
End Sub
Sub issueprize()

    
    fmequestions.Visible = False                                'on frmMain the frame that holds the questions is hidden to re-display the chipset map
    Unload frmQuestions                                         'frmquestion is what physically links to the database, it needs to be unloaded to reloadn correct table on next set of quuestions
    Unload frmbATTLE                                            'If a battle has taken place the form is unloaded because it is no longer neeeded
    cmdBegin.Visible = True                                     're-displays the begin button for next set of questions
    wrongtemp = 0                                               '''''''''''''''''''''
    correcttemp = 0                                             ''  Clean Results  ''
    currentquestion = 0                                         '''''''''''''''''''''
    frmWorldMap.picShip.Left = frmWorldMap.picShip.Left + 160   '
    frmWorldMap.tmrCheckLocation.Enabled = True
    cmdAn1.Visible = False
    cmdAn2.Visible = False
    cmdAn3.Visible = False
    cmdAn4.Visible = False
    ConfirmAnswer.Visible = False
    
    lblProgress.Caption = "Question: " & currentquestion & " \ 10"
    lblCorrect.Caption = "Correct: " & correcttemp & ""
    lblWrong.Caption = "Wrong: " & wrongtemp & ""
    lblCurrentLevel.Caption = "Current Level: " & currentlevel & ""
    txtUserAnswer.Text = ""
    
    loopcounter = 0                         ''''''''''''''''''''''''''''
        Do                                  ''  Clean Random Sequence ''
            a(loopcounter) = 0              ''''''''''''''''''''''''''''
            loopcounter = loopcounter + 1
        Loop Until loopcounter > 9
    loopcounter = 0
    Unload Me
End Sub

Private Sub ConfirmAnswer_Click()

'this code when clicked checks their answer, if in-correct, it informs them of the correct answer
'...if correct, simply display correct, either way it alters variables to monitor scores and progression

    If txtUserAnswer.Text <> "" Then
    totalq = totalq + 1                                         'Variable for total questions answered ever
            If txtUserAnswer.Text = txtFields(4).Text Then      'If chosen answer is same as database correct answer
                lblresult.Caption = "Correct!"                  'inform user that they got the question corecct
                correcttemp = correcttemp + 1                   'variable to tell them how many they they have got right in this set of ten questions
                correctall = correctall + 1                     'variable used in score at game over
                nextquestion                                    'calls sub that physically searches database for next question according to array of random numbers
                updatetext                                      'Puts the loaded question into the question feilds on frmMain
            Else
                lblresult.Caption = "Incorrect, it was: " & txtFields(4).Text & ""
                wrongtemp = wrongtemp + 1
                wrongall = wrongall + 1
                nextquestion
                updatetext
            End If
    Else
            lblresult.Caption = "Pick An Answer First!"         'If the user does not pick an answer before clicking confirm the are told to try again
    End If
    
End Sub

Private Sub Form_Keydown(KeyCode As Integer, Shift As Integer)

    'The idea for this technique was of a small game called samurai scroll
    '...made by "Jesse Acosta". however to use it i have had to completely re-code it from scratch
    '...this was because the program had many bugs in it, and errors in the structure
    '...plus i added functionality, to make the hero appear to walk and pass over floors other than grass.
    '...the plagiarism check would not pick this up, for i have re-written it all from scratch with
    '...my own variables, wording and structure, however i have told you anyway because i did not
    '...come up with the basic idea.
    
    FrmMain.fmequestions.Left = 0
    FrmMain.fmequestions.Top = 0
    FrmMain.fmequestions.Width = FrmMain.ScaleWidth
    FrmMain.fmequestions.Height = FrmMain.ScaleHeight
    
    'Changes Direction of Hero, then Moves if possibles
    
    Select Case KeyCode
        Case vbKeyUp
            Char_Face = 2                                       'Make hero look up
            Move_Hero                                           'Draw hero looking up
            Tile = Mid(AreaGrid(HeroY + 1 - 1), HeroX + 1, 1)   'Keep variable tidy for hero movement
    
    'the if below checks if the user can walk in that direction, for example, if they are in front of a tree they cannot walkthroug it
                
                If Tile = "X" Or Tile = "O" Or Tile = "6" Or Tile = "3" Or Tile = "U" Or Tile = "N" Then
                    HeroY = HeroY - 1
                    Checkforbridge
                    Draw_Position
                End If
        Case vbKeyDown
            Char_Face = 1
            Move_Hero
            Tile = Mid(AreaGrid(HeroY + 1 + 1), HeroX + 1, 1)
                If Tile = "X" Or Tile = "O" Or Tile = "6" Or Tile = "3" Or Tile = "U" Or Tile = "N" Then
                    HeroY = HeroY + 1
                    Checkforbridge
                    Draw_Position
                End If
        Case vbKeyLeft
            Char_Face = 3
            Move_Hero
            Tile = Mid(AreaGrid(HeroY + 1), HeroX + 1 - 1, 1)
                If Tile = "X" Or Tile = "O" Or Tile = "6" Or Tile = "3" Or Tile = "U" Or Tile = "N" Then
                    HeroX = HeroX - 1
                    Checkforbridge
                    Draw_Position
                End If
        Case vbKeyRight
            Char_Face = 4
            Move_Hero
            Tile = Mid(AreaGrid(HeroY + 1), HeroX + 1 + 1, 1)
                If Tile = "X" Or Tile = "O" Or Tile = "6" Or Tile = "3" Or Tile = "U" Or Tile = "N" Then
                    HeroX = HeroX + 1
                    Checkforbridge
                    Draw_Position
                End If
        Case Is = 13
            Init_Game
        Case vbKeySpace
            'Talks to someone if right in front of them
            Talk_Script
    End Select
'Am I at The Village? if so they leave current location and return to world map
'...i have not finished this code yet, maybe if i have time i might
            Tile = Mid(AreaGrid(HeroY + 1), HeroX + 1, 1)
            If Tile = "U" Then At_Village
End Sub

Private Sub tmrFirstdraw_Timer()
    
'Plays loading bar animation aswell aswell as drawing pics for the first time

    If loadbar4.X2 < 664 Then
        loadbar4.X2 = loadbar4.X2 + 8
    Else
    
            Tile = Mid(AreaGrid(HeroY + 1 + 1), HeroX + 1, 1)
                If Tile = "6" Or Tile = "3" Or Tile = "U" Or Tile = "N" Then
                    HeroY = HeroY + 1
                    Char_Face = 1
                    Draw_Position
                End If
            tmrFirstdraw.Enabled = False
            txtConverse.Width = FrmMain.ScaleWidth
            txtConverse.Height = FrmMain.ScaleHeight - 328
    End If
End Sub

Private Sub tmrupdate_Timer()
    lblresult.Caption = ""
    tmrupdate.Enabled = False
End Sub

Private Sub cmdAn1_Click()
    txtUserAnswer.Text = txtFields(0).Text
End Sub

Private Sub cmdAn2_Click()
    txtUserAnswer.Text = txtFields(1).Text
End Sub

Private Sub cmdAn3_Click()
    txtUserAnswer.Text = txtFields(2).Text
End Sub

Private Sub cmdAn4_Click()
    txtUserAnswer.Text = txtFields(3).Text
End Sub

Sub Checkforbridge()

'If the hero is going to pass over a floor other than grass then the heros chip...
'...set must be altered to account for the new background tile.
'Also when the hero moves the chip will change to make it look like...
'...they are walking, one step after another etc...

    Select Case Tile
    
    Case Is = "O"
        ImgUpG.Picture = onbridge1up.Picture
        ImgDownG.Picture = imgonbridgedown1.Picture
    Case Is = "3"
        ImgUpG.Picture = imgonbridge2up.Picture
        ImgDownG.Picture = mgonbridge2down.Picture
    Case Is = "X"
        ImgUpG.Picture = imgonbridge3up.Picture
        ImgDownG.Picture = imgonbridge3down.Picture
    Case Is = "6"
                If ImgDownG.Picture = imgdownleft.Picture Then
                    ImgDownG.Picture = down.Picture
                Else
                    ImgDownG.Picture = imgdownleft.Picture
                End If
                If ImgUpG.Picture = up.Picture Then
                    ImgUpG.Picture = imgupleft.Picture
                Else
                    ImgUpG.Picture = up.Picture
                End If
                If ImgLG.Picture = left1.Picture Then
                    ImgLG.Picture = left2.Picture
                Else
                    ImgLG.Picture = left1.Picture
                End If
                If ImgRG.Picture = right.Picture Then
                    ImgRG.Picture = right2.Picture
                Else
                    ImgRG.Picture = right.Picture
                End If
        
    End Select

End Sub

Private Sub Form_Load()
    frmWorldMap.tmrCheckLocation.Enabled = False
    frmWorldMap.picShip.Left = frmWorldMap.picShip.Left + 160
    Init_Game
End Sub


Sub Init_Game()

'this process finds the users current level/location on world map and draws a map for each one
'...the map information stored in an array is stored in the module "MapStore" in their own subs
'...once a map has been loaded into array, this sub calls the draw map funtion

FrmMain.Show
Found_Scroll = False

'V = Edge
'OceanEdgeNorth

Newline = Chr(13) + Chr(10)

Select Case currentlevel

Case 1          'General Knowledge
    LevelOne
Case 2          'Biology
    LevelTwo
Case 3          'Mathematics
    LevelThree
Case 4          'Physics
    LevelFour
Case 5          'Nature
    LevelFive
Case 6          'Geography
    LevelSix
Case 7          'Riddles
    LevelSeven
Case 8          'Media
    LevelEight
Case 9          'Music
    LevelNine
Case 10         'films
    LevelTen
End Select

    Draw_Position
End Sub

Sub Draw_Position()
'Draws map on form by getting string, using the len function to isolate a certain letter, takes that
'...letter and passes it through a case statement, the letter is matched to a picture, that picture is then
'...drawn according to the heros position on the map, stored in the X nad Y variable.
'...The picture is drawn onto the form using the paint picture command, this solves the problem of having to
'...use loads of image boxes.

Dim pass As Byte
'draw map at 10 tiles high by 22 tiles wide
'...the loop starts from a negative value to account for the hero being inset from 0,0

    For Y = -6 To 6 Step 1
        For X = -6 To 18 Step 1
        
    'A couple of IFs to define the edge of a map, if the hero walkss to the edge of the map, they might
    '...be able to see past the end of the array map defined in the module, so if a tile does not have a
    '...tile in the map array, then the sea tile is used, this is done via calling the Fill_in_with_water sub
            
            pass = 0
            If Y + HeroY + 0 < 1 Then Fill_in_with_water
            
            If X + HeroX + 0 < 1 Then Fill_in_with_water
            
            If X + HeroX + 0 > Len(AreaGrid(1)) Then Fill_in_with_water
            
            If Y + HeroY + 0 > 36 Then Fill_in_with_water
            
            If pass = 0 Then Tile = Mid(AreaGrid(Y + HeroY + 1), (X + HeroX + 1), 1)
            
            If X = 0 And Y = 0 Then GoTo Skip:
            
    'The bulky case statement applies a picture to any letter
            
        Select Case Tile

            Case Is = "0"
                PaintPicture ImgBB.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "1"
                PaintPicture ImgBl.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "2"
                PaintPicture imgBr.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "3"
                PaintPicture ImgBridge.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "4"
                PaintPicture ImgBush.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "5"
                PaintPicture ImgForest.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "6"
                PaintPicture ImgGrass.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "7"
                PaintPicture ImgGuy1.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "8"
                PaintPicture ImgGuy2.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "9"
                PaintPicture ImgGuy3.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "A"
                PaintPicture ImgGuy4.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "B"
                PaintPicture ImgGuy5.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "C"
                PaintPicture ImgIlb.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "D"
                PaintPicture ImgIlt.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "E"
                PaintPicture ImgIrb.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "F"
                PaintPicture ImgIrt.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "G"
                PaintPicture ImgL.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "H"
                PaintPicture ImgMount1.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "I"
                PaintPicture ImgMount2.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "J"
                PaintPicture ImgMount3.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "K"
                PaintPicture ImgMount4.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "L"
                PaintPicture ImgMountain.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "M"
                PaintPicture ImgR.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "N"
                PaintPicture ImgScroll.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "O"
                PaintPicture ImgSign.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "P"
                PaintPicture ImgSkull.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "Q"
                PaintPicture ImgTl.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "R"
                PaintPicture ImgTori.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "S"
                PaintPicture ImgTr.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "T"
                PaintPicture ImgTT.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "U"
                PaintPicture IMgVillage.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "V"
                PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "W"
                PaintPicture imgTreestop.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "X"
                PaintPicture imgbrifgenorth.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "Y"
                PaintPicture imgFrontWall.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "Z"
                PaintPicture imgWallback.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "a"
                PaintPicture imgWallLeft.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "b"
                PaintPicture imgStoneFloor.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "c"
                PaintPicture imgWallRight.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "d"
                PaintPicture imgwall.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "e"
                PaintPicture imgleftedgewall.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "f"
                PaintPicture imgwall3.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "g"
                PaintPicture imgWall4.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "h"
                PaintPicture imgwall5.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "i"
                PaintPicture imgdoors.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "j"
                PaintPicture imgwindow.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "k"
                PaintPicture imgtorch.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "l"
                PaintPicture imgEl.Picture, (X + 3) * 32, (Y + 3) * 32
            Case Is = "m"
                PaintPicture imgm.Picture, (X + 3) * 32, (Y + 3) * 32
        End Select
Skip:
        Next
    Next

Put_down_Character:
    Move_Hero           'the previous statement drew the grass, now the hero must be drawn, facing the right direction

End Sub
Sub Move_Hero()
'The placement of the charachter is simple, note the 3 * 32, this is because the tiles are 32 pixels wide,
'...and we want to put the hero to stand 3 tiles inset from top left (0,0)

    Select Case Char_Face
        Case Is = 1
                    PaintPicture ImgDownG.Picture, 3 * 32, 3 * 32
        Case Is = 2
                    PaintPicture ImgUpG.Picture, 3 * 32, 3 * 32
        Case Is = 3
                    PaintPicture ImgLG.Picture, 3 * 32, 3 * 32
        Case Is = 4
                    PaintPicture ImgRG.Picture, 3 * 32, 3 * 32
    End Select
    
    Look = Mid(AreaGrid(HeroY + 1), HeroX + 1, 1)
End Sub
Sub Fill_in_with_water()
    pass = 1
    PaintPicture ImgBlack.Picture, (X + 3) * 32, (Y + 3) * 32
End Sub

Sub Talk_Script()

'What happens when talking (MUST be underneath the person talking to)
'For example, in this case where X is the hero and Y is a talk target
'       .........Y.........
'       .........X.........

    Dim Choice As Byte
    Dim Look As String
    Dim Interect As Boolean
    Dim Speaker As String
    Look = Mid(AreaGrid(HeroY + 1 - 1), HeroX + 1, 1)
    
    Select Case Look
        Case Is = "7"
            Interect = True
            Speaker = "Bob: Greetings"
        Case Is = "8"
            Interect = True
            Speaker = "Homer: Good luck challenging the Question Master Sir."
        Case Is = "9"
            Interect = True
            Speaker = "Zack: Good Morning."
        Case Is = "A"
            Interect = True
            Speaker = "Shinji: Hi, welcome to the Thengal Province!"
        Case Is = "B"
            Interect = True
            Speaker = "The Question Master:"                        'Talking to this guy triggers the entire quiz section of the game
    
    'Message to user asking them if they want to start the quiz"
    
            EnterMessageText = "Do you wish to challenge me? 10 question according to your inteligence."
    'because i have made my own message boxes, i call a global sub depending on what type of question i am asking
    
            MeassageEnterLocation
        Case Is = "R"
            Interect = True
            Speaker = "Gaurd: Welcome To Xu Bay Sir!"
        Case Is = "O"
            Interect = True
            Speaker = "Zenal: Welcome to the area Sir!"
        Case Is = "i"
            Interect = True
            Speaker = "Thengal temple: LOCKED, Serol the mage came and killed the priest, the high priest will pay much for his capture"
            txtConverse.Text = txtConverse.Text + "" & Heroname & ": I must find Serol, he has gone mad!" + Newline + Newline
    End Select

    If Interect = True Then
        txtConverse.Text = txtConverse.Text + "" + Speaker + Newline + Newline
        If Look <> "i" Then txtConverse.Text = txtConverse.Text + "" & Heroname & ": OK, Thankyou." + Newline + Newline
    Else
        txtConverse.Text = txtConverse.Text + "" & Heroname & ": Erm, I seem to be talking to myself, not a good sign!" + Newline + Newline
    Interect = False
    End If
    txtConverse.SelStart = Len(txtConverse)
End Sub

Sub At_Village()
'temporary code, will be replaced with code allowing the user to quit the area
        txtConverse.Text = txtConverse.Text + "Region Master: Travel the area, find the question master and challenge him if you wish." + Newline + Newline
        txtConverse.SelStart = Len(txtConverse)
        txtConverse.Text = txtConverse.Text + "" & Heroname & ": Not yet, this is a difficult task." + Newline
        txtConverse.SelStart = Len(txtConverse)
End Sub

Sub setquestionsforgo()
        FrmMain.fmequestions.Visible = True             'the question boxes are on a frame, this maximises the frame to fill the screen
        randomisequestions                              'Calls the sub that randomises and loads questions
        currentquestion = 0                             're-set a few variables in case there is old data in them
        correcttemp = 0
        wrongtemp = 0
        frmWorldMap.tmrCheckLocation.Enabled = False    'stops frmWorlMap animation, whilst in quiz
End Sub

