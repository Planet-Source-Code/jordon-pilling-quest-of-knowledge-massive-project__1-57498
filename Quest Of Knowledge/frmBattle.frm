VERSION 5.00
Begin VB.Form frmBattle 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Combat"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   Picture         =   "frmBattle.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClaw 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9840
      Top             =   0
   End
   Begin VB.Timer tmrfirespell 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10320
      Top             =   0
   End
   Begin VB.Timer tmrATB 
      Interval        =   50
      Left            =   10800
      Top             =   0
   End
   Begin VB.Label lblDamage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbloption3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbloption2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Magic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblattack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgoption3 
      Height          =   360
      Left            =   120
      Picture         =   "frmBattle.frx":20A49
      Top             =   4200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image imgoption2 
      Height          =   360
      Left            =   120
      Picture         =   "frmBattle.frx":2288B
      Top             =   3720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image imgoption 
      Height          =   360
      Left            =   120
      Picture         =   "frmBattle.frx":246CD
      Top             =   3240
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image imgMenu 
      Height          =   1650
      Left            =   0
      Picture         =   "frmBattle.frx":2650F
      Top             =   3080
      Width           =   1875
   End
   Begin VB.Label lblEnemyStat 
      BackStyle       =   0  'Transparent
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
      Left            =   6480
      TabIndex        =   5
      Top             =   4320
      Width           =   5415
   End
   Begin VB.Image imgenemyATBFront 
      Height          =   105
      Left            =   2190
      Picture         =   "frmBattle.frx":306E1
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   45
   End
   Begin VB.Image imgprogress 
      Height          =   105
      Left            =   2190
      Picture         =   "frmBattle.frx":30779
      Stretch         =   -1  'True
      Top             =   3975
      Width           =   45
   End
   Begin VB.Image imgatbbar 
      Height          =   360
      Left            =   2040
      Picture         =   "frmBattle.frx":30811
      Top             =   3840
      Width           =   4440
   End
   Begin VB.Image imgAttack 
      Height          =   1185
      Left            =   4080
      Picture         =   "frmBattle.frx":30E55
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgclaw 
      Height          =   495
      Index           =   3
      Left            =   1440
      Picture         =   "frmBattle.frx":3110A
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image imgclaw 
      Height          =   495
      Index           =   2
      Left            =   960
      Picture         =   "frmBattle.frx":31804
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image imgclaw 
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "frmBattle.frx":31DEC
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image imgclaw 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmBattle.frx":3217B
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Image imgspell 
      Height          =   1440
      Left            =   3720
      Picture         =   "frmBattle.frx":32430
      Top             =   600
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   8
      Left            =   3840
      Picture         =   "frmBattle.frx":3329A
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   7
      Left            =   3360
      Picture         =   "frmBattle.frx":341FD
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   6
      Left            =   2880
      Picture         =   "frmBattle.frx":3511D
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   5
      Left            =   2400
      Picture         =   "frmBattle.frx":35FE4
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   4
      Left            =   1920
      Picture         =   "frmBattle.frx":36D66
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   3
      Left            =   1440
      Picture         =   "frmBattle.frx":379E4
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   2
      Left            =   960
      Picture         =   "frmBattle.frx":385F3
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "frmBattle.frx":392D1
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image fire 
      Height          =   500
      Index           =   0
      Left            =   0
      Picture         =   "frmBattle.frx":39FE0
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.Image Image8 
      Height          =   30
      Left            =   0
      Picture         =   "frmBattle.frx":3AE4A
      Stretch         =   -1  'True
      Top             =   3080
      Width           =   11970
   End
   Begin VB.Shape shpenemypointer 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   495
      Left            =   10680
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblhint 
      BackStyle       =   0  'Transparent
      Caption         =   "Let The Fight Begin..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   9975
   End
   Begin VB.Image Image6 
      Height          =   30
      Left            =   0
      Picture         =   "frmBattle.frx":3B3C4
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   13890
   End
   Begin VB.Image imgup 
      Height          =   360
      Left            =   9840
      Picture         =   "frmBattle.frx":3B93E
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image imgover 
      Height          =   360
      Left            =   9840
      Picture         =   "frmBattle.frx":3D780
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
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
      Left            =   6480
      TabIndex        =   0
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Image imgenemy1 
      Height          =   2280
      Left            =   10080
      Picture         =   "frmBattle.frx":3F5C2
      Top             =   480
      Width           =   1035
   End
   Begin VB.Image imgme 
      Height          =   1740
      Left            =   960
      Picture         =   "frmBattle.frx":40DF8
      Top             =   360
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2040
      Picture         =   "frmBattle.frx":424B0
      Top             =   4320
      Width           =   4440
   End
End
Attribute VB_Name = "frmbATTLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chooseenemy, attackwho As Boolean
Dim target As String
Dim fireanimation, clawanimation, damagedone As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub imgenemy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Line1.X2 = 8520
    Line1.Y2 = 1800
End Sub
Private Sub Form_Load()
    WholeProcessTrigger = True 'Stops world map cross hairs and everything whilst battle is going on
    hpcurrent = hp
    Select Case currentlevel
        Case 1
            currentenemymaxhp = 12
            currentenemyhp = 12
            currentenemyattack = 8
            currentenemywisdom = 9
        Case 2
            currentenemymaxhp = 22
            currentenemyhp = 22
            currentenemyattack = 14
            currentenemywisdom = 14
        Case 3
            currentenemymaxhp = 38
            currentenemyhp = 38
            currentenemyattack = 20
            currentenemywisdom = 20
        Case 4
            currentenemymaxhp = 58
            currentenemyhp = 58
            currentenemyattack = 29
            currentenemywisdom = 29
        Case 5
            currentenemymaxhp = 83
            currentenemyhp = 83
            currentenemyattack = 38
            currentenemywisdom = 38
    End Select
    
    frmWorldMap.tmrCheckLocation.Enabled = False
    Unload frmMessageBox
    Unload FrmMain
End Sub

Private Sub imgenemy1_Click()
    target = "" & currentenemyname & ""
    chooseenemy = False
    shpenemypointer.Visible = False
    imgprogress.Width = 75
    lblHint.Caption = "OK!"
    tmrClaw.Enabled = True
    attackwho = True
    killbattlemenu
End Sub

Private Sub imgenemy1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chooseenemy = True Then
        shpenemypointer.Visible = True
        shpenemypointer.Left = imgenemy1.Left + imgenemy1.Width / 2
        shpenemypointer.Top = imgenemy1.Top - shpenemypointer.Height
    End If
End Sub

Private Sub imgMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgoption.Picture = imgup.Picture
    imgoption2.Picture = imgup.Picture
    imgoption3.Picture = imgup.Picture
End Sub

Private Sub imgoption2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgoption2.Picture = imgover.Picture
End Sub

Private Sub imgoption3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgoption3.Picture = imgover.Picture
End Sub

Private Sub lblattack_Click()
    choosenemey
    killbattlemenu
End Sub

Private Sub lblattack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMenu.Visible = True Then
        imgoption.Picture = imgover.Picture
    End If
End Sub

Private Sub lbloption2_Click()

    If imgenemyATBFront.Width >= 4125 Then
        attackwho = False
        imgenemyATBFront.Width = 75
        lblHint.Caption = "" & currentenemyname & " Counters you're spell and diflects it back!"
    Else
        attackwho = True
        lblHint.Caption = "Fire Spell"
    End If
    
    imgprogress.Width = 75
    tmrfirespell.Enabled = True
    killbattlemenu
End Sub

Private Sub lbloption2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMenu.Visible = True Then
        imgoption2.Picture = imgover.Picture
    End If
End Sub

Private Sub lbloption3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgMenu.Visible = True Then
        imgoption3.Picture = imgover.Picture
    End If
End Sub

Private Sub tmrATB_Timer()
    If currentenemyhp <= 0 Then
        frmbATTLE.tmrATB.Enabled = False
        frmbATTLE.tmrClaw.Enabled = False
        frmbATTLE.tmrfirespell.Enabled = False
        WholeProcessTrigger = False 're-starts world map cross hairs and everything whilst battle is going on
        frmWorldMap.tmrCheckLocation.Enabled = True
        Unload frmMessageBox            'Close any idle message boxes
        Unload frmWorldMap
        frmWorldMap.Show
        SpeedOfShip = 0
        Unload frmbATTLE
    End If
    If hpcurrent <= 0 Then
        gameoverreason = "death"
        gameover        'if gamers hit points has been depleted then game over
    End If

    Unload frmMessageBox
    currentenemyskill = 16
    currentenemyname = "Question Master IIV"

If imgprogress.Width < 4125 Then
    imgprogress.Width = imgprogress.Width + skill * 4
    lblStat.Caption = "" & Heroname & "  HP: " & hpcurrent & " / " & hp & ""
Else
    battlemenu
End If

frmbATTLE.Show
frmbATTLE.SetFocus

If imgenemyATBFront.Width < 4125 Then
    imgenemyATBFront.Width = imgenemyATBFront.Width + currentenemyskill * 2
    lblEnemyStat.Caption = "" & currentenemyname & "  HP: " & currentenemyhp & " / " & currentenemymaxhp & ""
Else
    attackwho = False
    imgenemyATBFront.Width = 75
    lblHint.Caption = "The Question Master Attacks You With Claw Attack!"
    tmrClaw.Enabled = True
    killbattlemenu
End If

End Sub

Sub battlemenu()
    imgoption.Visible = True
    imgoption2.Visible = True
    imgoption3.Visible = True
    lblattack.Visible = True
    lbloption2.Visible = True
    lbloption3.Visible = True
End Sub
Sub killbattlemenu()
    imgoption.Visible = False
    imgoption2.Visible = False
    imgoption3.Visible = False
    lblattack.Visible = False
    lbloption2.Visible = False
    lbloption3.Visible = False
End Sub

Sub choosenemey()
    chooseenemy = True
    lblHint.Caption = "Click Enemy!"
End Sub

Sub bounceHealth()
        
        lblDamage.Caption = "" & damagedone & ""
        lblDamage.Visible = True
        If attackwho = True Then
            lblDamage.Top = 960
            lblDamage.Left = 10200
            floatdamage
        Else
            lblDamage.Top = 1080
            lblDamage.Left = 1080
            floatdamage 'Calls damage label animation
        End If
        
End Sub
Sub floatdamage()   ''''''Makes the amount of damage hover over the victims head in typical rpg fashion
Do
    Sleep 2
    lblDamage.Top = lblDamage.Top - 10
Loop Until lblDamage.Top <= 360

Do
    Sleep 2
    lblDamage.Top = lblDamage.Top + 10
Loop Until lblDamage.Top >= 1080

Do
    Sleep 2
    lblDamage.Top = lblDamage.Top - 10
Loop Until lblDamage.Top <= 400

Do
    Sleep 2
    lblDamage.Top = lblDamage.Top + 10
Loop Until lblDamage.Top >= 900

lblDamage.Visible = False

End Sub
Private Sub tmrClaw_Timer()    'The claw animation for an attack
If attackwho = True Then

        imgAttack.Left = 10080
        imgAttack.Top = 960
        imgAttack.Visible = True
        animateclaw

Else
        imgAttack.Left = 720
        imgAttack.Top = 600
        imgAttack.Visible = True
        animateclaw
End If

End Sub

Private Sub tmrfirespell_Timer()
If attackwho = True Then

        imgspell.Left = 9720
        imgspell.Top = 960
        imgspell.Visible = True
        animateflame

Else
        imgspell.Left = 960
        imgspell.Top = 720
        imgspell.Visible = True
        animateflame
End If

End Sub

Sub animateflame()

        Select Case fireanimation
        
            Case 0
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 1
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 2
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 3
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 4
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 5
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 6
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 7
                imgspell.Picture = fire(fireanimation)
                fireanimation = fireanimation + 1
                 
            Case 8
                imgspell.Picture = fire(fireanimation)
                fireanimation = 0
                imgspell.Visible = False
                tmrfirespell.Enabled = False
                imgenemy1.Visible = True
    Randomize
                If attackwho = True Then
                                    
                    damagedone = Int((wisdom * Rnd) * 2)
                    currentenemyhp = currentenemyhp - damagedone
                Else
                    damagedone = Int((currentenemywisdom * Rnd) * 2)
                    hpcurrent = hpcurrent - damagedone
                End If
                bounceHealth
        End Select

End Sub

Sub animateclaw()

           Select Case clawanimation
        
            Case 0
                imgAttack.Picture = imgclaw(0)
                clawanimation = clawanimation + 1
            Case 1
                imgAttack.Picture = imgclaw(1)
                clawanimation = clawanimation + 1
            Case 2
                imgAttack.Picture = imgclaw(2)
                clawanimation = clawanimation + 1
            Case 3
                imgAttack.Picture = imgclaw(3)
                clawanimation = clawanimation + 1
            Case 4
                imgAttack.Picture = imgclaw(3)
                clawanimation = clawanimation + 1
            Case 5
                imgAttack.Picture = imgclaw(2)
                clawanimation = clawanimation + 1
            Case 6
                imgAttack.Picture = imgclaw(1)
                clawanimation = clawanimation + 1
            Case 7
                imgAttack.Picture = imgclaw(1)
                clawanimation = clawanimation + 1
                tmrClaw.Enabled = False
                clawanimation = 0
                
        Randomize
                If attackwho = True Then
                    damagedone = attack * Rnd * 2
                    currentenemyhp = currentenemyhp - damagedone
                Else
                    damagedone = currentenemyattack * Rnd * 2
                    hpcurrent = hpcurrent - damagedone
                End If
                bounceHealth
                imgAttack.Visible = False

End Select

End Sub


