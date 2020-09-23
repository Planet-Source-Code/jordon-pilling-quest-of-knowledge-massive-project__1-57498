VERSION 5.00
Begin VB.Form frmOpen 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Choose location to save the profile settings..."
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   Picture         =   "open.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
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
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "savegame"
      Top             =   3000
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1290
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6855
   End
   Begin VB.Image imgokdown 
      Height          =   750
      Left            =   1440
      Picture         =   "open.frx":5642C
      Top             =   3960
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokupafter 
      Height          =   750
      Left            =   1440
      Picture         =   "open.frx":5BCB8
      Top             =   4680
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokup 
      Height          =   750
      Left            =   4800
      Picture         =   "open.frx":61544
      Top             =   2760
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub imgokup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgokup.Picture = imgokdown.Picture
    
End Sub

Private Sub imgokup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgokup.Picture = imgokupafter.Picture
    FileName = Dir1.Path & "\"
    nameoffile = Text1.Text
    frmmakenew.Text1.Text = "" & FileName & "" & nameoffile & ".SAV"
    frmmakenew.Show
    Unload Me
    
End Sub
