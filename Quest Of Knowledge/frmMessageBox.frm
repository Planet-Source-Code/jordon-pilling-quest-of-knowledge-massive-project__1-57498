VERSION 5.00
Begin VB.Form frmMessageBox 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Message..."
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMessageBox.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton optoption1 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton optOption2 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you really wish to quit?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
   Begin VB.Image imgCrit 
      Height          =   240
      Left            =   120
      Picture         =   "frmMessageBox.frx":28928
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgType 
      Height          =   480
      Left            =   360
      Picture         =   "frmMessageBox.frx":2956C
      Top             =   840
      Width           =   465
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LeftKey = False  'stops any world map movement acceleration
RightKey = False
UpKey = False
End Sub

Private Sub optOption1_Click()
    MBoxReturn = True
    Me.Hide
End Sub

Private Sub optOption2_Click()
    MBoxReturn = False
    Me.Hide
End Sub

