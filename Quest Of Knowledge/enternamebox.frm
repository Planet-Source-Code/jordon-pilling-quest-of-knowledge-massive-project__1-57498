VERSION 5.00
Begin VB.Form frmEnterName 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "+"
   ClientHeight    =   1920
   ClientLeft      =   2610
   ClientTop       =   3600
   ClientWidth     =   6540
   Icon            =   "enternamebox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "enternamebox.frx":0442
   ScaleHeight     =   1920
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   3720
      TabIndex        =   4
      Text            =   "18"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtenternameHero 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Text            =   "SparrowHawk"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtentername 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Text            =   "New Player"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hero Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image imgokdown 
      Height          =   750
      Left            =   240
      Picture         =   "enternamebox.frx":28D6A
      Top             =   4200
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokupafter 
      Height          =   750
      Left            =   600
      Picture         =   "enternamebox.frx":2E5F6
      Top             =   2400
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokup 
      Height          =   750
      Left            =   4200
      Picture         =   "enternamebox.frx":33E82
      Top             =   960
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmEnterName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim length As Integer
Dim message As String

Private Sub imgokup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgokup.Picture = imgokdown.Picture
    
End Sub

Private Sub imgokup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Checks name length is between 3 and 35 charachters, age is a sensible age between 3 and 100
'... and that all feild are valid, not blank.
'A precise message will tell the user exactly what is wrong, and even put an '&' if more than one is wrong
    
    imgokup.Picture = imgokupafter.Picture
    age = txtAge.Text
    If age = "" Or age > 100 Or age < 3 Then
        EnterMessageText = "You're age is not valid, cannot be null, and must be between 3 and 100."
        OKMessage
        Exit Sub
    End If
    
    Playername = txtentername.Text

    length = Len(Playername)

    If length <= 35 And length >= 3 Then
    Else
        message = " Player Name"
    End If
    age = txtAge.Text
    
    Heroname = txtenternameHero.Text

    length = Len(Heroname)

    If length <= 35 And length >= 3 Then
    Else
        If message = "" Then
        message = "" & message & " Hero-name"
        Else
        message = "" & message & " and Hero-name"
        End If
    End If
    If message <> "" Then
        EnterMessageText = "The entered information for" & message & ", is not valid, they must be between 3 and 35 characters in length."
        OKMessage
        message = ""
    Else
        frmmakenew.lblname.Caption = "" & Heroname & ""
        Unload Me
    End If
End Sub

