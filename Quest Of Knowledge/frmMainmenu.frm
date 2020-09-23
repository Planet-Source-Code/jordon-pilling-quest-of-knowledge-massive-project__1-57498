VERSION 5.00
Begin VB.Form frmMainmenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainmenu.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   5055
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      Begin VB.PictureBox cmdquit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmMainmenu.frx":D9838
         ScaleHeight     =   750
         ScaleWidth      =   2250
         TabIndex        =   7
         Top             =   3480
         Width           =   2280
      End
      Begin VB.PictureBox cmdAdmin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmMainmenu.frx":DF0C4
         ScaleHeight     =   750
         ScaleWidth      =   2250
         TabIndex        =   6
         Top             =   1800
         Width           =   2280
      End
      Begin VB.PictureBox cmdhelp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmMainmenu.frx":E4950
         ScaleHeight     =   750
         ScaleWidth      =   2250
         TabIndex        =   5
         Top             =   2640
         Width           =   2280
      End
      Begin VB.PictureBox cmdoption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmMainmenu.frx":EA1DC
         ScaleHeight     =   750
         ScaleWidth      =   2250
         TabIndex        =   4
         Top             =   960
         Width           =   2280
      End
      Begin VB.PictureBox cmdStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   120
         Picture         =   "frmMainmenu.frx":EFA68
         ScaleHeight     =   750
         ScaleWidth      =   2250
         TabIndex        =   3
         Top             =   120
         Width           =   2280
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   1560
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   10935
         TabIndex        =   1
         Top             =   4560
         Width           =   10935
         Begin VB.Label lblComment 
            BackStyle       =   0  'Transparent
            Caption         =   "Choose An Option..."
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
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   10815
         End
      End
      Begin VB.Image quitup 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":F52F4
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image quitdown 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":FAB80
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image helpup 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":10040C
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image helpdown 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":105C98
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image adminup 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":10B524
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image admindown 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":110DB0
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image loadup 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":11663C
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image loaddown 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":11BEC8
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image newup 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":121754
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Image newdown 
         Height          =   750
         Left            =   2520
         Picture         =   "frmMainmenu.frx":126FE0
         Top             =   3480
         Visible         =   0   'False
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmMainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed1 As Integer
Dim counter As Integer
Dim fIn As Boolean
Dim flag As Integer
Dim red As Integer
Dim posofshape, flag1 As Boolean
Dim fontsizeint As Integer
Dim loop1 As Integer
Dim stage As Integer

Private Sub cmdAdmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
    cmdAdmin.Picture = admindown.Picture
    lblComment.Caption = "The Admin section is protected by password, allows editing of scores and options."
End Sub

Private Sub cmdhelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
    cmdhelp.Picture = helpdown.Picture
End Sub

Private Sub cmdOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
    cmdoption.Picture = loaddown.Picture
    lblComment.Caption = "THIS FUNCTION IS NOT AVILABLE IN THIS VERSION"
End Sub

Private Sub cmdQuit_Click()
    MeassageQuit
End Sub

Private Sub cmdquit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
    cmdQuit.Picture = quitdown.Picture
End Sub

Private Sub cmdStart_Click()

    frmOpen.Show
    Timer1.Enabled = False
    Unload Me

End Sub

Private Sub cmdAdmin_Click()
    frmAdminLogin.Show
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
    cmdStart.Picture = newdown.Picture
    lblComment.Caption = "Click here to start a new game, You will be asked where you want to save you're settings and progress."
End Sub

Private Sub Form_Load()
rolls = 10
speed1 = 100
stage = 0
Dim L_nRun As Integer
Dim lngReturnResult As Long
On Error Resume Next
counter = 0
flag = 0
roll = False
fIn = True
    For L_nRun = 0 To 1000
            With G_dStar(L_nRun)
                .nX = Int(Rnd * frmMainmenu.Picture1.ScaleWidth) + 100
                .nY = Int(Rnd * frmMainmenu.Picture1.ScaleHeight) + 100
                Select Case Int(Rnd * 9) + 1
                    Case 1
                        .nSpeed = 30
                        .nColor = &HFFFFFF
                    Case 2, 3, 4
                        .nSpeed = 15
                        .nColor = &H808080
                    Case 5, 6, 7, 8, 9
                        .nSpeed = 10
                        .nColor = &H404040
                End Select
            End With
        Next
End Sub

Private Sub Picture1_Click()
    Unload frmAdminLogin
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cleanButtons
End Sub

Private Sub Timer1_Timer()

    For L_nRun = 0 To 1000
        With G_dStar(L_nRun)
           frmMainmenu.Picture1.PSet (.nX, .nY), &H0&
           .nX = .nX - .nSpeed
           If .nX < 5 Then .nX = frmMainmenu.Picture1.ScaleWidth
           frmMainmenu.Picture1.PSet (.nX, .nY), .nColor
        End With
    Next

End Sub

Sub cleanButtons()
    cmdStart.Picture = newup.Picture
    cmdoption.Picture = loadup.Picture
    cmdAdmin.Picture = adminup.Picture
    cmdhelp.Picture = helpup.Picture
    cmdQuit.Picture = quitup.Picture
End Sub
