VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Quest Of Knowledge"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoadBar 
      Interval        =   50
      Left            =   5280
      Top             =   720
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Image imgLoadingBar 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   240
      Picture         =   "frmSplash.frx":28928
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   15
   End
   Begin VB.Image imgloadingback 
      Height          =   135
      Left            =   240
      Picture         =   "frmSplash.frx":289D2
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   6090
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim proceed As Boolean


Private Sub tmrLoadBar_Timer()

If proceed = False Then
    If imgLoadingBar.Width < imgloadingback.Width Then
        imgLoadingBar.Width = imgLoadingBar.Width + 152
    Else
        lblComment.Caption = "STAND BY"
        tmrLoadBar.Interval = 2000
        proceed = True
        GoTo 19:
    End If
End If
    
    If proceed = True Then
        Unload Me
        frmIntro.Show
    End If
19:
End Sub
