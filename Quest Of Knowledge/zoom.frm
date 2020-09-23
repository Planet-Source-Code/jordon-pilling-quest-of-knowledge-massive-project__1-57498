VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   345
   ClientTop       =   0
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   60
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrMoveBanners 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.TextBox lblIntroText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   240
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By J.Pilling @ The Phoenix Studios   V.1.1 2004"
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
      Height          =   255
      Left            =   -120
      TabIndex        =   2
      Top             =   5760
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Phoenix Studios"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   11175
   End
   Begin VB.Shape Banner 
      BackColor       =   &H006B0500&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   0
      Left            =   0
      Top             =   -1300
      Width           =   11220
   End
   Begin VB.Shape Banner 
      BackColor       =   &H006B0500&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   1
      Left            =   0
      Top             =   6000
      Width           =   11220
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim red1, animationstep As Integer
Dim colour, path1, fullpath, AllText1, LineOfText1 As String
Dim phasebanners As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'this sub is heavily based on stuff i found on the internet once, i no longer have the url
'...but the explode form idea is not mine, the fading text and stuff later on is my idea though

Sub ExplodeForm(frm As Form, Optional ByVal lNumSteps As Long = 25, _
    Optional ByVal lStepDuration As Long)
    Dim formleft As Single, formtop As Single
    Dim sngHeight As Single, sngWidth As Single
    Dim sngNewHeight As Single, sngNewWidth As Single
    Dim sngHeightStep As Single, sngWidthStep As Single
    Dim iStep As Long
    Dim formx, formy As Single
        
    On Error Resume Next
    If frm.WindowState <> vbNormal Then Exit Sub
    
    formleft = frm.Left
    formtop = frm.Top
    sngHeight = frm.Height
    sngWidth = frm.Width

    sngHeightStep = sngHeight / lNumSteps
    sngWidthStep = sngWidth / lNumSteps
        For iStep = 1 To lNumSteps
            sngNewHeight = sngNewHeight + sngHeightStep
            sngNewWidth = sngNewWidth + sngWidthStep
            frm.Move formleft + (sngWidth - sngNewWidth) / 2, _
                formtop + (sngHeight - sngNewHeight) / 2, sngNewWidth, sngNewHeight
              frm.Visible = True
            frm.Refresh
            Sleep lStepDuration
        Next
    frm.Move formleft, formtop, sngWidth, sngHeight
    Timer1.Enabled = False
    Banner(0).Visible = True
    Banner(1).Visible = True
    tmrMoveBanners.Enabled = True
End Sub

Private Sub Form_Click()
    killIntro                   'Allows the user to skip the intro
End Sub

Private Sub Form_Load()
'sets animation to ready

    path1 = App.Path
    fullpath = "" & path1 & "\intro2.txt"

Label1.ForeColor = RGB(0, 0, 0)
red1 = 0
Timer1.Enabled = True

End Sub

Private Sub Label1_Click()
    killIntro               'Allows the user to skip the intro
End Sub

Private Sub lblIntroText_Click()
killIntro                   'Allows the user to skip the intro
End Sub

Private Sub picBackground_Click()
    killIntro               'Allows the user to skip the intro
End Sub
Public Sub killIntro()
    lblIntroText.Visible = False
    Label1.Visible = False
    Timer3.Enabled = False
    frmMainmenu.Show
    Unload Me
End Sub
Private Sub Timer1_Timer()
    ExplodeForm Me, 80, 10  'Starts explode form function
End Sub

Private Sub Timer3_Timer()

'this code is all done by myself, it simply displays words and fades them to black,
'...and then swaps words andd fades them back to white.
'...each case handles a different word, then passes onto next word

Select Case animationstep

Case 0
        Label1.Visible = True
        
            If red1 >= 0 And red1 < 255 And phasing = False Then
                red1 = red1 + 5
            Else
                phasing = True
            End If
        
        If red1 > 0 And red1 <= 255 And phasing = True Then
        red1 = red1 - 5
            If red1 = 5 Then
                phasing = False
                Label1.Caption = "Proudly Presents"
                animationstep = 1
            End If
        End If
        Label1.ForeColor = RGB(red1, red1, red1)
Case 1
            Label1.Visible = True
            
            If red1 >= 0 And red1 < 255 And phasing = False Then
                red1 = red1 + 5
            Else
                phasing = True
            End If
        
        If red1 > 0 And red1 <= 255 And phasing = True Then
        red1 = red1 - 5
            If red1 = 5 Then
                phasing = False
                Label1.Caption = "The Quest Of Knowledge"
                Label1.FontBold = True
                animationstep = 2
            End If
        End If
        Label1.ForeColor = RGB(red1, red1, red1)
Case 2
        Label1.Visible = True
        
            If red1 >= 0 And red1 < 255 And phasing = False Then
                red1 = red1 + 5
                Label1.ForeColor = RGB(red1, 0, 0)
            Else
                phasing = True
                
                animationstep = 3
            End If
Case 3
        If red1 > 0 And red1 <= 255 And phasing = True Then
            red1 = red1 - 5
            Label1.ForeColor = RGB(red1, 0, 0)
        Else
            animationstep = 5
            Label1.Visible = False
            On Error GoTo 26:
            Open fullpath For Input As #1
            Exit Sub
26:
            lblIntroText.Visible = True
        End If


Case 5
    'this steps reads from the filem opened in the last case, it displays
    '...an intro letter by letter.
    
                    Timer3.Interval = 50
        On Error GoTo 23:
                    lblIntroText.Visible = True
                    Line Input #1, LineOfText1
                    AllText1 = AllText1 & LineOfText1
                    lblIntroText.Text = AllText1
                    red1 = 255
        Exit Sub
        
Case 6
23:
    Close #1
    Timer3.Interval = 20
        If Banner(0).FillColor <> vbBlack Then
           Banner(0).FillColor = RGB(red1, red1, red1)
            Label4.Top = Label4.Top - 20
        Else
                phasebanners = True
        End If
End Select
End Sub

Private Sub tmrMoveBanners_Timer()

'when the form is dis[played to blue banners will grow in making the form lookd cinematic
'...i just did this for aesthetic reasons

If phasebanners = False Then
    If Banner(1).Top > 4800 Then
        Banner(1).Top = Banner(1).Top - 100
        Banner(0).Top = Banner(0).Top + 100
    End If
Else
    If Banner(1).Top < frmIntro.ScaleHeight Then
    tmrMoveBanners.Interval = 50
        Banner(1).Top = Banner(1).Top + 50
        Banner(0).Top = Banner(0).Top - 50
        On Error GoTo 24:
        lblIntroText.Height = lblIntroText.Height - 150
24:
    Else
            Timer3.Enabled = False
            frmMainmenu.Show
            Unload Me
    End If
End If
End Sub
