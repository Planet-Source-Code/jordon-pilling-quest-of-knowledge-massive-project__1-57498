VERSION 5.00
Begin VB.Form frmOpenPic 
   BorderStyle     =   0  'None
   Caption         =   "Load Portrait..."
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7200
   Icon            =   "openpic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "openpic.frx":08CA
   ScaleHeight     =   3675
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
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
      Height          =   1080
      Left            =   120
      Pattern         =   "*.bmp;*.wmf;*.jpg"
      TabIndex        =   2
      Top             =   2280
      Width           =   4335
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
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
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H006B0500&
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
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4335
   End
   Begin VB.Image cmdQuit 
      Height          =   750
      Left            =   4680
      Picture         =   "openpic.frx":56CF6
      Top             =   2760
      Width           =   2250
   End
   Begin VB.Image imgokupafter 
      Height          =   750
      Left            =   840
      Picture         =   "openpic.frx":5C582
      Top             =   5640
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokdown 
      Height          =   750
      Left            =   480
      Picture         =   "openpic.frx":61E0E
      Top             =   7440
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   -360
      Top             =   -240
      Width           =   495
   End
   Begin VB.Image imgAccessed 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "frmOpenPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdQuit.Picture = imgokdown.Picture
End Sub

Private Sub cmdQuit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdQuit.Picture = imgokupafter.Picture
If imgAccessed.Picture = Image1.Picture Then
    EnterMessageText = "No Picture was selected!"
    OKMessage
    Unload Me
Else
    frmmakenew.imgPortrait.Picture = frmOpenPic.imgAccessed.Picture
    Unload Me
End If
End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    
    'This statement links together the file list box and the directory...
    '...list box so the files shown match together
    
End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    
    'This hooks the two objects together,Drive & Dir,so that the...
    '...directory list box will show folders of the selected drive
    
End Sub

Private Sub File1_Click()
    SelectedFile = File1.Path & "\" & File1.FileName
    imgAccessed.Picture = LoadPicture(SelectedFile)
    
    'this will string the selections together to form one variable which...
    '...will hold the filename, this is then accessed and displayed into theimage box
End Sub

Private Sub Form_Load()
Drive1.Drive = "c:"
Dir1.Path = "c:\"

End Sub
