VERSION 5.00
Begin VB.Form frmWorldMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "The Quest Of Knowledge - World Map"
   ClientHeight    =   10740
   ClientLeft      =   -30
   ClientTop       =   -150
   ClientWidth     =   14670
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   DrawWidth       =   4
   Icon            =   "frmgameboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmgameboard.frx":0442
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   978
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCheckLocation 
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   960
      Picture         =   "frmgameboard.frx":240486
      ScaleHeight     =   0.01
      ScaleMode       =   0  'User
      ScaleWidth      =   0.01
      TabIndex        =   11
      Top             =   1200
      Width           =   15
   End
   Begin VB.Shape shpBlankout 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbly 
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblx 
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Heading:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   3135
   End
   Begin VB.Line shpcompass 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   136
      X2              =   136
      Y1              =   0
      Y2              =   24
   End
   Begin VB.Shape shpCircle 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   495
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgPortrait 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line shpHorizontal 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   64
      X2              =   80
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line shpvert 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   72
      X2              =   72
      Y1              =   0
      Y2              =   16
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Village"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape home 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   3120
      Shape           =   1  'Square
      Top             =   2040
      Width           =   150
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "The Lost Library"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tronx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   8520
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Shelandor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13200
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Xu Bay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Es-AssA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sh-Hi-Na"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ten Prison"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Zanori Tessa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MoRR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   7080
      Width           =   735
   End
   Begin VB.Shape Level9 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   10800
      Shape           =   1  'Square
      Top             =   1800
      Width           =   150
   End
   Begin VB.Shape Level10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   7680
      Shape           =   1  'Square
      Top             =   8400
      Width           =   150
   End
   Begin VB.Shape Level6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   4200
      Shape           =   1  'Square
      Top             =   8640
      Width           =   150
   End
   Begin VB.Shape Level8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   13680
      Shape           =   1  'Square
      Top             =   1560
      Width           =   150
   End
   Begin VB.Shape Level7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   10080
      Shape           =   1  'Square
      Top             =   5520
      Width           =   150
   End
   Begin VB.Shape Level4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   1560
      Shape           =   1  'Square
      Top             =   7800
      Width           =   150
   End
   Begin VB.Shape Level3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   1200
      Shape           =   1  'Square
      Top             =   2880
      Width           =   150
   End
   Begin VB.Shape Level2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   5880
      Shape           =   1  'Square
      Top             =   720
      Width           =   150
   End
   Begin VB.Shape Level5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   6480
      Shape           =   1  'Square
      Top             =   7320
      Width           =   150
   End
   Begin VB.Shape Level1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   3720
      Shape           =   1  'Square
      Top             =   5640
      Width           =   150
   End
   Begin VB.Image imgshipright 
      Height          =   255
      Left            =   0
      Picture         =   "frmgameboard.frx":24294E
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgshipleft 
      Height          =   240
      Left            =   0
      Picture         =   "frmgameboard.frx":244E05
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgshipdown 
      Height          =   345
      Left            =   120
      Picture         =   "frmgameboard.frx":2472B1
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgshipup 
      Height          =   300
      Left            =   120
      Picture         =   "frmgameboard.frx":2497B9
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblThengal 
      BackStyle       =   0  'Transparent
      Caption         =   "Thengal Province"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frmWorldMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentquestion As Integer
Dim moveback As Boolean

Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Const Cherry = 3.14159              'Pi
Const GoPower = 0.1                 'Acceleration
Const TurningSpeed = 15             'Rotation speed
Const pointerlinelength = 20        'Size of white pointer line
Const DelayTimer = 25               'Milliseconds per frame

Dim mlngTimer As Long               'Timer to control ship loop
Dim AngleOfShip As Single           'Angle of the ship
Dim ShipHeading As Single           'Direction of ship
Dim Pointerx As Single              'X - CO-ORD of ship
Dim Pointery As Single              'Y - CO-ORD of ship

Private Sub Form_Load()
    imgPortrait.Picture = frmmakenew.imgPortrait.Picture
    Pointerx = 200
    Pointery = 200
    mlngTimer = GetTickCount()
    WholeProcessTrigger = True
    frmWorldMap.Show
    WholeProcessTrigger = False
    frmHelp.Show vbModal
    frmWorldMap.picShip.SetFocus

    
        Do While WholeProcessTrigger = False

        If mlngTimer + DelayTimer <= GetTickCount() Then
            mlngTimer = GetTickCount()
            Physics                     'update ships location
            DrawShip
        End If
            DoEvents
    Loop
   
End Sub


Private Sub picShip_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft And Not RightKey Then LeftKey = True
    If KeyCode = vbKeyRight And Not LeftKey Then RightKey = True
    If KeyCode = vbKeyUp Then UpKey = True

End Sub
Private Sub picShip_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyLeft Then LeftKey = False
    If KeyCode = vbKeyRight Then RightKey = False
    If KeyCode = vbKeyUp Then UpKey = False

End Sub

Private Sub quit_Click()
    MeassageQuit
End Sub

Private Sub tmrCheckLocation_Timer()
    locationcheck
End Sub

Private Sub Physics()

'The trigonometry involved in this next bit is really complex, and i dont really understand it myself
'...i got help of some guy in a VB chatroom at http://www.andreavb.com/mychat/index.php3

Dim sngXComp As Single
Dim sngYComp As Single
Dim i As Integer
    
    If RightKey = True Then
        AngleOfShip = AngleOfShip + TurningSpeed * Cherry / 180
    End If

    If LeftKey = True Then
        AngleOfShip = AngleOfShip - TurningSpeed * Cherry / 180
    End If
    
    If UpKey = True Then
        sngXComp = SpeedOfShip * Sin(ShipHeading) + GoPower * Sin(AngleOfShip)
        sngYComp = SpeedOfShip * Cos(ShipHeading) + GoPower * Cos(AngleOfShip)
        lblHeading.Caption = "Heading: " & AngleOfShip & ""
        SpeedOfShip = Sqr(sngXComp ^ 2 + sngYComp ^ 2)
        lblSpeed = "Speed: " & SpeedOfShip * 10 & "Mph"
        If sngYComp > 0 Then ShipHeading = Atn(sngXComp / sngYComp)
        If sngYComp < 0 Then ShipHeading = Atn(sngXComp / sngYComp) + Cherry
    End If
    
    'Mathematical bit
    Pointerx = Pointerx + SpeedOfShip * Sin(ShipHeading)
    Pointery = Pointery - SpeedOfShip * Cos(ShipHeading)
    
    'If ship goes of side of form, make it re-appear on opposite side
    If Pointerx > frmWorldMap.ScaleWidth Then Pointerx = 0
    If Pointery > frmWorldMap.ScaleHeight Then Pointery = 0
    If Pointerx < 0 Then Pointerx = frmWorldMap.ScaleWidth
    If Pointery < 0 Then Pointery = frmWorldMap.ScaleHeight

End Sub

Private Sub DrawShip()

    picShip.Top = Pointery
    picShip.Left = Pointerx
    
    lblx.Caption = "Longitude: " & Pointerx & ""
    lbly.Caption = "Latitude: " & Pointery & ""
    intX1 = Pointerx + pointerlinelength * Sin(AngleOfShip)
    intY1 = Pointery - pointerlinelength * Cos(AngleOfShip)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    WholeProcessTrigger = False
End Sub


