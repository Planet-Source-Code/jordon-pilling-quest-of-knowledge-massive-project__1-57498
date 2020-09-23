VERSION 5.00
Begin VB.Form frmAdminLogin 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Administrator Login..."
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAdminLogin.frx":0000
   ScaleHeight     =   1905
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtentername 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   1560
      MousePointer    =   12  'No Drop
      TabIndex        =   1
      Text            =   "Administrator"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "w"
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name:"
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
      Height          =   495
      Left            =   -600
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image imgokup 
      Height          =   750
      Left            =   4200
      Picture         =   "frmAdminLogin.frx":28928
      Top             =   960
      Width           =   2250
   End
   Begin VB.Image imgokupafter 
      Height          =   750
      Left            =   720
      Picture         =   "frmAdminLogin.frx":2E1B4
      Top             =   3240
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgokdown 
      Height          =   750
      Left            =   720
      Picture         =   "frmAdminLogin.frx":33A40
      Top             =   2520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Height          =   495
      Left            =   -600
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim enteredpassword, username, applicationpath As String
Dim passwordarray(255) As String

Private Sub Form_Load()

End Sub

Private Sub imgokup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgokup.Picture = imgokdown.Picture
End Sub

Private Sub imgokup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    open_and_decrypt
End Sub


Private Sub txtpassword_Keydown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then open_and_decrypt
End Sub

Sub open_and_decrypt()

imgokup.Picture = imgokupafter.Picture

    On Error GoTo filenotfound
       
    decrypt_and_compare
    
    If decryptedpassword = txtpassword.Text Then
        frmAdministrator.Show
        frmAdminLogin.Hide
        frmMainmenu.Hide
    Else
        EnterMessageText = "Incorrect Password, Please re-try"
        OKMessage
    End If
    
    Exit Sub

filenotfound:
    EnterMessageText = "You're Password file could not be Loaded, it has been replaced with default password, see your licence for this info"
    OKMessage
    Wrap$ = Chr$(13) + Chr$(10)
    Open applicationpath & "\Savedata.txt" For Output As #2
        Print #2, "216" & Wrap$ & "222" & Wrap$ & "232" & Wrap$ & "208" & Wrap$ & "216" & Wrap$ & "222" & Wrap$ & "228" & Wrap$ & "210" & Wrap$ & "202" & Wrap$ & "220" & Wrap$ & "" 'this is the default password
    Close #2

End Sub

Sub decrypt_and_compare()
Dim temp1, temp2, decryptedletter, passwordlength As String
'''''''''''''''''''''''''''''''''''''''''''''''''
'encypting the password is the easy part        '
'reversing the process however is going to be   '
'more tricky                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''

temp1 = enteredpassword

    applicationpath = App.Path
    FileNumber = FreeFile
    counter2 = 0
    
   'read in the sets of numbers line by line into an array
    
    Open applicationpath & "\Savedata.txt" For Input As #1
    Do
        Input #1, passwordarray(counter2)
        counter2 = counter2 + 1
        Loop Until EOF(1)
    Close #1
    
passwordlength = counter2 - 1 'how many loops where needed to read in password
counter2 = 0
decryptedpassword = ""
On Error GoTo passworderror:

'convert numbers (stored in array) into charachters and assemble into a string
'...thus re-constructing the password
Do
    temp2 = passwordarray(counter2)
    temp2 = temp2 / 2
    decryptedletter = Chr(temp2)
    decryptedpassword = "" & decryptedpassword & "" & decryptedletter & ""
    counter2 = counter2 + 1

Loop Until counter2 = passwordlength
Exit Sub

passworderror:
    EnterMessageText = "Password Error, password file is present but corrupt, password file will be deleted, and restored on next login"
    OKMessage
'erase corrupt password file
    Kill applicationpath & "\Savedata.txt"
    EnterMessageText = "The Quest Of Knowledge has deleted the corrupt password file for you, Please login with default password."
    OKMessage
    
End Sub
