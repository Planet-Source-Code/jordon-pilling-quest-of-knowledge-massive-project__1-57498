VERSION 5.00
Begin VB.Form frmGameOver 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameOver.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackground 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      FillColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   11460
      TabIndex        =   0
      Top             =   1400
      Width           =   11460
      Begin VB.Timer tmrsaveinfotodatabase 
         Interval        =   500
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer tmranimate 
         Interval        =   1
         Left            =   120
         Top             =   120
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   5895
         Left            =   0
         TabIndex        =   1
         Top             =   6000
         Width           =   11250
      End
   End
   Begin VB.Data dtascorelink 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\College\VB\Assignment 3\Quest Of Knowledge\Questions.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "scores"
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Image imgToMenu 
      Height          =   750
      Left            =   8880
      Picture         =   "frmGameOver.frx":124FC2
      Top             =   120
      Width           =   2250
   End
   Begin VB.Image imgtomenudown 
      Height          =   750
      Left            =   6000
      Picture         =   "frmGameOver.frx":1278F9
      Top             =   4200
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgtomenuupafter 
      Height          =   750
      Left            =   5880
      Picture         =   "frmGameOver.frx":129F94
      Top             =   5040
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label lblage 
      BackStyle       =   0  'Transparent
      DataField       =   "Age"
      DataSource      =   "dtascorelink"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblscore 
      BackStyle       =   0  'Transparent
      DataField       =   "Score"
      DataSource      =   "dtascorelink"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblname 
      BackStyle       =   0  'Transparent
      DataField       =   "Name"
      DataSource      =   "dtascorelink"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblsummary 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct / Incorrect:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   11055
   End
   Begin VB.Label lbltotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Answered Questions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   11055
   End
   Begin VB.Label lbloutcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Outcome:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   11055
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim red1 As Integer
Dim colour As String

Private Sub Form_Load()
colour = vbGreen
Randomize
red1 = 0
Dim L_nRun As Integer
On Error GoTo 20

'save score to database
    apppath = App.Path
    dtascorelink.DatabaseName = apppath & "\db1.mdb"
    dtascorelink.RecordSource = "scores"



On Error Resume Next
    For L_nRun = 0 To 150
            With G_dStar(L_nRun)
                .nY = Int(Rnd * frmGameOver.picBackground.ScaleHeight) + 100
                .nX = Int(Rnd * frmGameOver.picBackground.ScaleWidth) + 100
                Select Case Int(Rnd * 9) + 1
                    Case 1
                        .nSpeed = 50
                        .nColor = colour
                    Case 2, 3, 4
                        .nSpeed = 20
                        .nColor = colour
                    Case 5, 6, 7, 8, 9
                        .nSpeed = 5
                        .nColor = colour
                End Select
            End With
        Next
        
Exit Sub

20:
    EnterMessageText = "Score Database could not be found, youre scores will not be saved!"
    OKMessage
End Sub


Private Sub imgToMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgToMenu.Picture = imgtomenudown.Picture
End Sub

Private Sub imgToMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgToMenu.Picture = imgtomenuupafter.Picture
    frmMainmenu.Show
    Unload Me
End Sub

Private Sub tmrsaveinfotodatabase_Timer()
    frmGameOver.Show
    
    dtascorelink.Recordset.AddNew
    dtascorelink.Recordset.update
    dtascorelink.Recordset.MoveLast

    If gameoverreason = "death" Then lbloutcome.Caption = "Outcome: You Perished In Battle"
    If gameoverreason = "donequestions" Then lbloutcome.Caption = "Outcome: Congratulations, You defeated the elder question master!."
    lbltotal.Caption = "Answered Questions: " & totalq & ""
    lblsummary.Caption = "Correct / Incorrect: " & correctall & "/" & wrongall & ""
    lblage.Caption = "" & age & ""
    score = ((totalq - wrongall) * currentlevel) * diffnum
    lblscore.Caption = "" & score & ""
    lblName.Caption = "" & Playername & ""
    
    tmrsaveinfotodatabase.Enabled = False
End Sub

Private Sub tmranimate_Timer()
Dim x1, y1 As Integer
Randomize
    For L_nRun = 0 To 150
        With G_dStar(L_nRun)
            frmGameOver.picBackground.PSet (.nX, .nY), colour
            .nY = .nY + .nSpeed
            variableline = Int(1700 * Rnd)
            frmGameOver.picBackground.PSet (.nX, .nY - variableline), vbBlack
                If .nY > frmGameOver.ScaleHeight Then .nY = 0
                    frmGameOver.picBackground.PSet (.nX, .nY), colour
                    x1 = Int((6015 * Rnd) + 4920)
                    y1 = Int((11220 * Rnd) + 0)
                    frmGameOver.picBackground.PSet (y1, x1), vbBlack
        End With
    Next
End Sub

