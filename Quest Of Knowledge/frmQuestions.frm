VERSION 5.00
Begin VB.Form frmQuestions 
   BorderStyle     =   0  'None
   Caption         =   "Easy"
   ClientHeight    =   3210
   ClientLeft      =   1050
   ClientTop       =   0
   ClientWidth     =   7590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7590
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Next Question via random pattern"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdClearArray 
      Caption         =   "Clear Array"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview Question Order."
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Answer1"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   3375
   End
   Begin VB.Data dtaEasyLink 
      Align           =   2  'Align Bottom
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon Contol\My Documents\HND Year 1\Visual Programming\Assignment 3\Quest Of Knowledge\Questions.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Easy1"
      Top             =   2865
      Width           =   7590
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Question Number"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   2100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Question"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   10
      Top             =   1785
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Correct Answer"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   1455
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Answer4"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   1140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Answer3"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   825
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Answer2"
      DataSource      =   "dtaEasyLink"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   495
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Question Number:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Question:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Correct Answer:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Answer4:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Answer3:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Answer2:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Answer1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim duplicate As Integer
Dim firstloopflag As Integer

Private Sub cmdsearch_Click()
    nextquestion
End Sub

Private Sub Command1_Click()
    randomisequestions
End Sub

Private Sub cmdClearay_Click()
loopcounter = 0
Do
    a(loopcounter) = 0
    loopcounter = loopcounter + 1
Loop Until loopcounter > 9

End Sub

Private Sub Form_Load()
On Error GoTo 20:
    apppath = App.Path
    dtaEasyLink.DatabaseName = apppath & "\db1.mdb"
    dtaEasyLink.RecordSource = "" & Diff & "" & currentlevel & ""
Exit Sub

20:
    EnterMessageText = "Database could not be found, Please re-install"
    OKMessage
End Sub


