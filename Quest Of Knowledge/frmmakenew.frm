VERSION 5.00
Begin VB.Form frmmakenew 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Quest Of Knowledge - Making A New Character"
   ClientHeight    =   9165
   ClientLeft      =   870
   ClientTop       =   1740
   ClientWidth     =   13440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmakenew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmakenew.frx":0442
   ScaleHeight     =   9165
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer update 
      Interval        =   1
      Left            =   1680
      Top             =   8520
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
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
      Height          =   585
      Left            =   -4000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   12
      Top             =   7440
      Width           =   2220
   End
   Begin VB.Timer tmrtidy 
      Interval        =   3
      Left            =   11640
      Top             =   3120
   End
   Begin VB.Frame fmeMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   0
      TabIndex        =   15
      Top             =   1400
      Width           =   3615
      Begin VB.Label lblHint 
         BackStyle       =   0  'Transparent
         Caption         =   "HINT:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgportraitup 
         Height          =   750
         Left            =   720
         Picture         =   "frmmakenew.frx":125406
         Top             =   3360
         Width           =   2250
      End
      Begin VB.Image imgstartup 
         Height          =   750
         Left            =   720
         Picture         =   "frmmakenew.frx":127C13
         Top             =   2520
         Width           =   2250
      End
      Begin VB.Image cmdroll 
         Height          =   750
         Left            =   720
         Picture         =   "frmmakenew.frx":12A2FC
         Top             =   1680
         Width           =   2250
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stats Left:"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   1920
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Image imgtomenuup 
         Height          =   750
         Left            =   720
         Picture         =   "frmmakenew.frx":12C826
         Top             =   4200
         Width           =   2250
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   3360
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Label lblHoverhint 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Charachter Maker"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   21
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label lblDifficulty 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   3960
      Width           =   3855
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   5160
      TabIndex        =   18
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Height          =   495
      Left            =   2640
      Shape           =   2  'Oval
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblWisdom 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Wisdom:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   8280
      Width           =   855
   End
   Begin VB.Image imgwisdomdown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":12F15D
      Top             =   8475
      Width           =   405
   End
   Begin VB.Image imgwisdomup 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":12F498
      Top             =   8160
      Width           =   405
   End
   Begin VB.Image imgrightdown 
      Height          =   405
      Left            =   8640
      Picture         =   "frmmakenew.frx":12F7C7
      Top             =   -240
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgrightafter 
      Height          =   405
      Left            =   8520
      Picture         =   "frmmakenew.frx":12FB1A
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   2760
      Picture         =   "frmmakenew.frx":12FE46
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   405
   End
   Begin VB.Image downafter 
      Height          =   285
      Left            =   8280
      Picture         =   "frmmakenew.frx":130172
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image upafter 
      Height          =   285
      Left            =   8400
      Picture         =   "frmmakenew.frx":1304AD
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgdepdown 
      Height          =   285
      Left            =   8160
      Picture         =   "frmmakenew.frx":1307DC
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgdepup 
      Height          =   285
      Left            =   8160
      Picture         =   "frmmakenew.frx":130B33
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image skilup 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":130E9C
      Top             =   2760
      Width           =   405
   End
   Begin VB.Image skilldown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":1311CB
      Top             =   3075
      Width           =   405
   End
   Begin VB.Image luckup 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":131506
      Top             =   3840
      Width           =   405
   End
   Begin VB.Image luckdown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":131835
      Top             =   4155
      Width           =   405
   End
   Begin VB.Image attup 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":131B70
      Top             =   4920
      Width           =   405
   End
   Begin VB.Image attdown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":131E9F
      Top             =   5235
      Width           =   405
   End
   Begin VB.Image defup 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":1321DA
      Top             =   6000
      Width           =   405
   End
   Begin VB.Image defdown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":132509
      Top             =   6315
      Width           =   405
   End
   Begin VB.Image hpDown 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":132844
      Top             =   7395
      Width           =   405
   End
   Begin VB.Image hpUp 
      Height          =   285
      Left            =   2760
      Picture         =   "frmmakenew.frx":132B7F
      Top             =   7080
      Width           =   405
   End
   Begin VB.Label lblSavepath 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Path:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Image imgblank 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgrollupafter 
      Height          =   750
      Left            =   10560
      Picture         =   "frmmakenew.frx":132EAE
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgrolldown 
      Height          =   750
      Left            =   10320
      Picture         =   "frmmakenew.frx":1353D8
      Top             =   240
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgtomenuupafter 
      Height          =   750
      Left            =   8040
      Picture         =   "frmmakenew.frx":137600
      Top             =   360
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgtomenudown 
      Height          =   750
      Left            =   8040
      Picture         =   "frmmakenew.frx":139F37
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgcontupafter 
      Height          =   750
      Left            =   7680
      Picture         =   "frmmakenew.frx":13C5D2
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgcontdown 
      Height          =   750
      Left            =   7680
      Picture         =   "frmmakenew.frx":13EEC4
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgportraitupafter 
      Height          =   750
      Left            =   7680
      Picture         =   "frmmakenew.frx":141549
      Top             =   240
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgportraitdown 
      Height          =   750
      Left            =   7320
      Picture         =   "frmmakenew.frx":143D56
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label lblhpvalue 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lbldefencevalue 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblattackvalue 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblLuckValue 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblskillvalue 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblStartHpheader 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Constitution:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblLuckheader 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligence:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lbldefence 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defence:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblattackheader 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Attack:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblSkillheader 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Skill:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblnameheader 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image imgportrait 
      BorderStyle     =   1  'Fixed Single
      Height          =   2820
      Left            =   120
      Picture         =   "frmmakenew.frx":14624D
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2310
   End
   Begin VB.Shape header 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmmakenew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim messagetext As String

'This is a very chunky form due to all the button mouse over affects and stat gaurds
'Not much i can do about it without using a directX component that does mouse up button effects

Private Sub attdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    attdown.Picture = imgdepdown.Picture
End Sub

Private Sub attdown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - this determines how much damage you do to your enemy."
    Shape1.Top = attdown.Top - 100
    Shape1.Left = attdown.Left - 110
End Sub

Private Sub attdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If attack > 3 Then
    attack = attack - 1
    availablestats = availablestats + 1
    lblattackvalue.Caption = "" & attack & ""
    lblLeft.Caption = "" & availablestats & ""
End If
Shape1.BorderColor = vbBlue
attdown.Picture = downafter.Picture
End Sub

Private Sub attup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    attup.Picture = imgdepup.Picture
End Sub

Private Sub attup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - this determines how much damage you do to your enemy."
    Shape1.Top = attup.Top - 100
    Shape1.Left = attup.Left - 110
End Sub

Private Sub attup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If availablestats > 0 Then
    If attack < 18 Then
        attack = attack + 1
        availablestats = availablestats - 1
        lblattackvalue.Caption = "" & attack & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
End If
Shape1.BorderColor = vbBlue
attup.Picture = upafter.Picture
End Sub

Private Sub cmdroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdroll.Picture = imgrolldown.Picture
    
End Sub

Private Sub cmdroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Generates stats via 8d8, then deducts minimums to stats for you."
End Sub

Private Sub cmdroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'when the user clicks roll, a number is generated and then spread over stats to fulfill
'...minimum values for each stat. the user is then free to apply the remaining stats to
'...what ever they want.

    cmdroll.Picture = imgrollupafter.Picture

        availablestats = Int((60 * Rnd) + 26)
        rolls = rolls - 1
        lblLeft.Caption = "" & availablestats & ""
        hp = 8
        availablestats = availablestats - 8
        
        attack = 3
        availablestats = availablestats - 3
        
        defence = 7
        availablestats = availablestats - 7
        
        skill = 3
        availablestats = availablestats - 3
                
        luck = 1
        availablestats = availablestats - 1
        
        wisdom = 3
        availablestats = availablestats - 3
        
        lblskillvalue.Caption = "" & skill & ""
        lblLuckValue.Caption = "" & luck & ""
        lblhpvalue.Caption = "" & hp & ""
        lblattackvalue.Caption = "" & attack & ""
        lbldefencevalue.Caption = "" & defence & ""
        lblWisdom.Caption = "" & wisdom & ""
        
        lblLeft.Caption = "" & availablestats & ""
End Sub

Private Sub defdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    defdown.Picture = imgdepdown.Picture
End Sub

Private Sub defdown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - your defence is compared to your enemies attack to see if a hit occurs."
    Shape1.Top = defdown.Top - 100
    Shape1.Left = defdown.Left - 110
End Sub

Private Sub defdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If defence > 7 Then
        defence = defence - 1
        availablestats = availablestats + 1
        lbldefencevalue.Caption = "" & defence & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
    Shape1.BorderColor = vbBlue
    defdown.Picture = downafter.Picture
End Sub

Private Sub defup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    defup.Picture = imgdepup.Picture
End Sub

Private Sub defup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - your defence is compared to your enemies attack to see if a hit occurs."
    Shape1.Top = defup.Top - 100
    Shape1.Left = defup.Left - 110
End Sub

Private Sub defup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If availablestats > 0 Then
        If defence < 18 Then
            defence = defence + 1
            availablestats = availablestats - 1
            lbldefencevalue.Caption = "" & defence & ""
            lblLeft.Caption = "" & availablestats & ""
        End If
    End If
    Shape1.BorderColor = vbBlue
    defup.Picture = upafter.Picture
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHoverhint.Left = X + 300
lblHoverhint.Top = Y
End Sub

Private Sub hpDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    hpDown.Picture = imgdepdown.Picture
End Sub

Private Sub hpDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - if you run out of health you will die, be wise!"
    Shape1.Top = hpDown.Top - 100
    Shape1.Left = hpDown.Left - 110
End Sub

Private Sub hpDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbBlue
    If hp > 11 Then
        hp = hp - 1
        availablestats = availablestats + 1
        lblhpvalue.Caption = "" & hp & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
    hpDown.Picture = downafter.Picture
End Sub

Private Sub hpUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    hpUp.Picture = imgdepup.Picture
End Sub

Private Sub hpUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - if you run out of health you will die, be wise!"
    Shape1.Top = hpUp.Top - 100
    Shape1.Left = hpUp.Left - 110
End Sub

Private Sub hpUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbBlue
    hpUp.Picture = upafter.Picture
    If availablestats > 0 Then
        If hp < 22 Then
            hp = hp + 1
            availablestats = availablestats - 1
            lblhpvalue.Caption = "" & hp & ""
            lblLeft.Caption = "" & availablestats & ""
        End If
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    hpUp.Picture = imgdepup.Picture
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = imgrightdown.Picture
    Shape1.BorderColor = vbRed
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Click here to enter you're name and the name for you're hero."
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.Picture = imgrightafter.Picture
    Shape1.BorderColor = vbBlue
    frmEnterName.Show vbModal, Me
End Sub

Private Sub imgportrait_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: You can change this picture by clicking PORTRAIT on the side menu."
End Sub

Private Sub imgportraitup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgportraitup.Picture = imgportraitdown.Picture
End Sub

Private Sub imgportraitup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Click here to load a portrait pic from you're HDD."
End Sub

Private Sub imgportraitup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgportraitup.Picture = imgportraitupafter.Picture
    frmOpenPic.Show
End Sub

Private Sub imgstartup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgstartup.Picture = imgcontdown.Picture
End Sub

Private Sub imgstartup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Start The Quest, you will be told if any Data is missing."
End Sub

Private Sub imgstartup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Before the user can start the game they must have entered valid information, the following code
'...checks stats have been roled and all applied, aswell as a portrait being chosen and finally
'...They have entered a name.


    If availablestats > 0 Then
        EnterMessageText = "You still have stats in youre stat pool, please apply them to a stat."
        GoTo 20:
    End If
    
    imgstartup.Picture = imgcontupafter.Picture
    If frmmakenew.imgportrait.Picture = Empty Then
        EnterMessageText = "Please choose A portrait before you continue!"
        GoTo 20:
    End If
    If lblskillvalue.Caption = "" Then
        EnterMessageText = "Please Roll Your STATS By Clicking Roll!"
        GoTo 20:
    End If
    If lblName.Caption = "" Then
        EnterMessageText = "Please enter your name before you continue!"
        GoTo 20:
    End If

    

frmWorldMap.Show
update.Enabled = False
Unload frmMessageBox
Unload Me
Exit Sub



20:
    If EnterMessageText <> "" Then
        OKMessage
    End If
End Sub


Private Sub imgtomenuup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgtomenuup.Picture = imgtomenudown.Picture
End Sub

Private Sub imgtomenuup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Return to the main menu discards all settings and stats."
End Sub

Private Sub imgtomenuup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgtomenuup.Picture = imgtomenuupafter.Picture
    frmMainmenu.Show
    Unload Me
End Sub


Private Sub imgwisdomdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    imgwisdomdown.Picture = imgdepdown.Picture
End Sub

Private Sub imgwisdomdown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lblHoverhint.Caption = "HINT: Combat - You're Wsidom will decide how much magic damage you will do."
    Shape1.Top = imgwisdomdown.Top - 100
    Shape1.Left = imgwisdomdown.Left - 110
End Sub

Private Sub imgwisdomdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If wisdom > 3 Then
        wisdom = wisdom - 1
        availablestats = availablestats + 1
        lblWisdom.Caption = "" & wisdom & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
    Shape1.BorderColor = vbBlue
    imgwisdomdown.Picture = downafter.Picture
End Sub

Private Sub imgwisdomup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    imgwisdomup.Picture = imgdepup.Picture
End Sub

Private Sub imgwisdomup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Combat - You're Wsidom will decide how much magic damage you will do."
    Shape1.Top = imgwisdomup.Top - 100
    Shape1.Left = imgwisdomup.Left - 110
End Sub

Private Sub imgwisdomup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbBlue
    imgwisdomup.Picture = upafter.Picture
    If availablestats > 0 Then
        If wisdom < 18 Then
            wisdom = wisdom + 1
            availablestats = availablestats - 1
            lblWisdom.Caption = "" & wisdom & ""
            lblLeft.Caption = "" & availablestats & ""
        End If
    End If
End Sub

Private Sub luckdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    luckdown.Picture = imgdepdown.Picture
End Sub

Private Sub luckdown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: The higher you raise the Intelligence the harder the questions shall be."
    Shape1.Top = luckdown.Top - 100
    Shape1.Left = luckdown.Left - 110
End Sub

Private Sub luckdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If luck > 1 Then
        luck = luck - 1
        availablestats = availablestats + 1
        lblLuckValue.Caption = "" & luck & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
    Shape1.BorderColor = vbBlue
    luckdown.Picture = downafter.Picture
End Sub

Private Sub luckup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    luckup.Picture = imgdepup.Picture
End Sub

Private Sub luckup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: The higher you raise the Intelligence the harder the questions shall be."
    Shape1.Top = luckup.Top - 100
    Shape1.Left = luckup.Left - 110
End Sub

Private Sub luckup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If availablestats > 0 Then
        If luck < 18 Then
            luck = luck + 1
            availablestats = availablestats - 1
            lblLuckValue.Caption = "" & luck & ""
            lblLeft.Caption = "" & availablestats & ""
        End If
    End If
    Shape1.BorderColor = vbBlue
    luckup.Picture = upafter.Picture
End Sub

Private Sub skilldown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.BorderColor = vbRed
    skilldown.Picture = imgdepdown.Picture
End Sub

Private Sub skilldown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Top = skilldown.Top - 100
    Shape1.Left = skilldown.Left - 110
    lblHoverhint.Caption = "HINT: Determines you ability to attack the enemy and actually hit."
End Sub

Private Sub skilldown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If skill > 3 Then
        skill = skill - 1
        availablestats = availablestats + 1
        lblskillvalue.Caption = "" & skill & ""
        lblLeft.Caption = "" & availablestats & ""
    End If
    Shape1.BorderColor = vbBlue
    skilldown.Picture = downafter.Picture
End Sub

Private Sub skilup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    skilup.Picture = imgdepup.Picture
    Shape1.BorderColor = vbRed
End Sub

Private Sub skilup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHoverhint.Caption = "HINT: Determines you ability to attack the enemy and actually hit."
    Shape1.Top = skilup.Top - 100
    Shape1.Left = skilup.Left - 110
End Sub

Private Sub skilup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If availablestats > 0 Then
        If skill < 18 Then
            skill = skill + 1
            availablestats = availablestats - 1
            lblskillvalue.Caption = "" & skill & ""
            lblLeft.Caption = "" & availablestats & ""
        End If
    End If
    Shape1.BorderColor = vbBlue
    skilup.Picture = upafter.Picture
End Sub

Private Sub tmrtidy_Timer()

' if form is loaded on different resolutions, the form should still stretch to fit screen
'...this code simply puts things where they should be.

    Text1.Width = frmmakenew.ScaleWidth
    Text1.Left = 0
    Text1.Top = frmmakenew.ScaleHeight - Text1.Height
    header.Width = frmmakenew.ScaleWidth
    fmeMenu.Left = frmmakenew.ScaleWidth - fmeMenu.Width
    fmeMenu.Height = frmmakenew.ScaleHeight - header.Height
    lblSavepath.Top = frmmakenew.ScaleHeight - lblSavepath.Height - Text1.Height
    tmrtidy.Enabled = False
    
End Sub

Private Sub update_Timer()

'stores difficulty in numerical and word form, numerical for calculating score
'...word form for searching database for correct record scource

    If luck <= 8 Then
        lblDifficulty.Caption = "Difficulty: Easy"
        Diff = "Easy"
        diffnum = 1
    End If
    If luck < 15 And luck > 8 Then
        lblDifficulty.Caption = "Difficulty: Medium"
        Diff = "Medium"
        diffnum = 2
    End If
    If luck >= 15 Then
        lblDifficulty.Caption = "Difficulty: Hard"
        Diff = "Hard"
        diffnum = 3
    End If

End Sub
