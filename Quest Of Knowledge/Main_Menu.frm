VERSION 5.00
Begin VB.Form Main_Menu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Slayer - Main Menu"
   ClientHeight    =   5190
   ClientLeft      =   915
   ClientTop       =   2175
   ClientWidth     =   9885
   Icon            =   "Main_Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   4680
   End
   Begin VB.Image imgexitupafter 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":0442
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgexitdown 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":2A6E
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgexitup 
      Height          =   750
      Left            =   7440
      Picture         =   "Main_Menu.frx":4DA0
      Top             =   4320
      Width           =   2250
   End
   Begin VB.Image imgenternameupafter 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":73CC
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgenternamedown 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":9E76
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgenternameup 
      Height          =   750
      Left            =   7440
      Picture         =   "Main_Menu.frx":C720
      Top             =   3480
      Width           =   2250
   End
   Begin VB.Image lblstartupafter 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":F1CA
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image lblstartdown 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":118B3
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgstartup 
      Height          =   750
      Left            =   7440
      Picture         =   "Main_Menu.frx":13CC7
      Top             =   2640
      Width           =   2250
   End
   Begin VB.Image imgloadupafter 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":163B0
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgloaddown 
      Height          =   750
      Left            =   0
      Picture         =   "Main_Menu.frx":18E68
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image Imgloadover 
      Height          =   750
      Left            =   7440
      Picture         =   "Main_Menu.frx":1B734
      Top             =   1800
      Width           =   2250
   End
   Begin VB.Label lblshowname 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image imgfighter 
      Height          =   4305
      Left            =   240
      Picture         =   "Main_Menu.frx":1E1EC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2700
   End
   Begin VB.Label lblcurrentplayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Player:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblslayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Slayer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   8280
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblthe 
      BackStyle       =   0  'Transparent
      Caption         =   "The"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgsword 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1485
   End
End
Attribute VB_Name = "Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim typecode As Integer
Dim reply As Integer



Private Sub Form_Load()

    lblshowname.Caption = heroname

End Sub

Private Sub imgenternameup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgenternameup.Picture = imgenternamedown.Picture

End Sub

Private Sub imgenternameup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgenternameup.Picture = imgenternameupafter.Picture
    
    enternamebox.Show
    Focus = Text1

End Sub

Private Sub imgexitup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgexitup.Picture = imgexitdown.Picture
    
End Sub

Private Sub imgexitup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgexitup.Picture = imgexitupafter.Picture
    
    quitverify.Show

End Sub

Private Sub Imgloadover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Imgloadover.Picture = imgloaddown

End Sub

Private Sub Imgloadover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Imgloadover.Picture = imgloadupafter
    frmsavingstuff.Show
    Main_Menu.Hide

End Sub


Private Sub imgstartup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgstartup.Picture = lblstartdown.Picture

End Sub

Private Sub imgstartup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgstartup.Picture = lblstartupafter.Picture
    
    newgameornot = True
        
    If lblshowname.Caption = "" Then

        MsgBox "Please Enter Your Name Before You Continue!", , "Enter Name"

    Else

        frmmakenew.Show
        Main_Menu.Hide
        frmmakenew.lblname.Caption = heroname

    End If

End Sub

Private Sub Timer1_Timer()

    lblshowname.Caption = heroname

End Sub
