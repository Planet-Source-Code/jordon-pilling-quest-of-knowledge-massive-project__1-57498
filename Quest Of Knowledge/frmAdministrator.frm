VERSION 5.00
Begin VB.Form frmAdministrator 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   Picture         =   "frmAdministrator.frx":0000
   ScaleHeight     =   5655
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   300
      Left            =   9360
      TabIndex        =   24
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Frame fmeOptions 
      BackColor       =   &H006B0500&
      Caption         =   "Change Password"
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
      Height          =   2655
      Left            =   5640
      TabIndex        =   14
      Top             =   960
      Width           =   5535
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   300
         Left            =   3840
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtConfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1560
         PasswordChar    =   "w"
         TabIndex        =   17
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtNew 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1560
         PasswordChar    =   "w"
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtOld 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1560
         PasswordChar    =   "w"
         TabIndex        =   15
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm New:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fmeChangepassword 
      BackColor       =   &H006B0500&
      Caption         =   "Options"
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
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   5175
      Begin VB.CheckBox Check1 
         BackColor       =   &H006B0500&
         Caption         =   "World Map Cross Hairs On/Off"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006B0500&
      Caption         =   "Scores stored in central database"
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
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5175
      Begin VB.Data dtalink 
         Appearance      =   0  'Flat
         Caption         =   "Scores, Please Browse"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Documents and Settings\Avalon Alpha\My Documents\HND Year 1\Visual Programming\Assignment 3\Quest Of Knowledge\Questions.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "scores"
         Top             =   1800
         Width           =   3840
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Save"
         Height          =   300
         Left            =   3840
         TabIndex        =   1
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox cmbSort 
         Height          =   315
         ItemData        =   "frmAdministrator.frx":D36F
         Left            =   120
         List            =   "frmAdministrator.frx":D371
         TabIndex        =   23
         Text            =   "Sort"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   2880
         TabIndex        =   2
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1920
         TabIndex        =   3
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Slot"
         DataSource      =   "dtalink"
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
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Age"
         DataSource      =   "dtalink"
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
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Name"
         DataSource      =   "dtalink"
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
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Score"
         DataSource      =   "dtalink"
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
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slot:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempstring, tempstring2, encryptedpassword As String
Dim passwordlength As Integer

Private Sub cmbSort_Click()
'A simple set of SQL commands designed to sort the database in different orders

Select Case cmbSort.ListIndex

Case 0
    EnterMessageText = "Sort Database by score order?"
    Yes_No
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then
            dtalink.RecordSource = "SELECT * FROM Scores ORDER BY Score"
            dtalink.Refresh
        End If
Case 1
    EnterMessageText = "Sort Database by NAME order?"
    Yes_No
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then
            dtalink.RecordSource = "SELECT * FROM Scores ORDER BY Name"
            dtalink.Refresh
        End If
Case 2
    EnterMessageText = "Sort Database by AGE order?"
    Yes_No
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then
            dtalink.RecordSource = "SELECT * FROM Scores ORDER BY Age"
            dtalink.Refresh
        End If
End Select
End Sub

Private Sub cmdAdd_Click()
    dtalink.Recordset.AddNew
    dtalink.Recordset.update
    dtalink.Recordset.MoveLast
End Sub

Private Sub cmdBack_Click()
    EnterMessageText = "Are you sure you want to quit to main menu?"
    Yes_No
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then
            frmMainmenu.Show
            Unload Me
        End If
End Sub

Private Sub cmdDelete_Click()
'Allows the user to delete a score record, once they have confirmed their action

    EnterMessageText = "Are you sure you want to Delete this record?"
    Yes_No
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then
            If dtalink.Recordset.EOF = False And dtalink.Recordset.BOF = False Then
                dtalink.Recordset.Delete
                dtalink.Recordset.MoveNext
                MsgBox "The record has been deleted"
            End If
        End If

End Sub

Private Sub cmdUpdate_Click()
    dtalink.Recordset.update
End Sub

Private Sub cmdchange_Click()
  On Error GoTo filenotfound
    
    applicationpath = App.Path
    FileNumber = FreeFile
     
    Open applicationpath & "\Savedata.txt" For Input As #FileNumber
        Input #FileNumber, enteredpassword
    Close #FileNumber
    
    If decryptedpassword = txtOld(4).Text Then
        If txtNew(5).Text = txtConfirm(6).Text Then
            encrypt_password_and_write
                Open applicationpath & "\Savedata.txt" For Output As #FileNumber
                    Print #FileNumber, encryptedpassword
                Close #FileNumber
            EnterMessageText = "You're password has been changed successfully."
            OKMessage
       Else
            EnterMessageText = "New passwords do not match, please check and try again."
            OKMessage
        End If
     Else
        EnterMessageText = "Incorrect Password, Please re-enter youre old password."
        OKMessage
     End If
    
    Exit Sub

filenotfound:
    EnterMessageText = "You're Password file could not be Loaded, it has been replaced with default password, see your licence for this info"
    OKMessage
    Open applicationpath & "\Savedata.txt" For Output As #FileNumber
        Print #FileNumber, "lothlorien" 'this is the default password
    Close #FileNumber

End Sub

Private Sub Form_Load()
    apppath = App.Path
    dtalink.DatabaseName = apppath & "\db1.mdb"
    dtalink.RecordSource = "scores"
    cmbSort.AddItem "Score"
    cmbSort.AddItem "Name"
    cmbSort.AddItem "Age"
End Sub

Sub encrypt_password_and_write()

'This sub modifies the new password into a pattern of numbers
'...this prevent users from opening the file and simply reading
'...the file in notepad, now all they will get is a load of meaningless numbers

'The encrytption is simple
'A) Split password into its indervidual charachters
'B) Convert these charachters into their ascii values
'C) Multply the ascii value by 2 so people who can read
'...ascii, wont recognise the numbers
Wrap$ = Chr$(13) + Chr$(10)
tempstring = txtConfirm(6).Text
passwordlength = Len(txtConfirm(6).Text)
counter2 = 1
Do
    tempstring2 = Mid(tempstring, counter2, 1)
    encryptedpassword = "" & encryptedpassword & "" & Asc(tempstring2) * 2 & "" & Wrap$ & ""
    counter2 = counter2 + 1
Loop Until counter2 > passwordlength

End Sub

