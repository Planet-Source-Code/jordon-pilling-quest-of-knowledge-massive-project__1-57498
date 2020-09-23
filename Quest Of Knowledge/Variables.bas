Attribute VB_Name = "Variables"
Option Explicit
'Animation Variables
    Global phasing, formwidth, formwidth1, formwidth2, animation1 As Integer 'Variable to control splash screen zoom on effect
    Global reply As Integer                     'Used for msgbox replies when my custom message box system cannot serve
    Global intX1, intY1 As Integer
    Global WholeProcessTrigger As Boolean       '''''''''''''''''''''''''
    Global SpeedOfShip As Single                'Speed of the ship      '
    Global LeftKey As Boolean                   'Left Cursor Trigger    '
    Global RightKey As Boolean                  'Right Cursor Trigger   '
    Global UpKey As Boolean                     'Up Cursor Trigger      '
                                                '''''''''''''''''''''''''
'System Variables                               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global savepath, numbers As String          'Will be used in save game function                                     '
    Global apppath As String                    'location of app for database location                                  '
    Global decryptedpassword As String          'Holds password once it has been de-crypted from a file                 '
    Global gameoverreason As String             'holds reason for game over, usually dead, or all questions completed   '
    Global firstloopflag As Integer             'used to control the stages of a loop through the program               '
    Global b As Integer                         'short for bufferm this will hold a temporary value for stuff           '

'MessageBox Variables                           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global EnterMessageText As String           'Message box text to be displayed                                       '
    Global MBoxReturn As Boolean                'What result the messagebox returns, (on a standard msgbox yes = 6 etc) '
                                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'players stats
    Global availablestats, rolls, luck, defence As Integer          '''''''''''''''''''''''''''''''''
    Global skill, hp, attack, wisdom, counter2, diffnum As Integer  'Stats to hold players values   '
    Global Diff, Playername, Heroname, age, score As String         '''''''''''''''''''''''''''''''''
                                                                                        
'Question Randomisation Variables               '''''''''''''''''''''''''''''''''''''''''''''
    Global loopcounter, a(9), buffer As Integer 'Array for question order and loop counters '
                                                '''''''''''''''''''''''''''''''''''''''''''''
'Enemy Battle Stats
    Global currentenemyname, currentenemyskill, currentenemyhp, currentenemymaxhp As Integer

'Hero Battle Stats                              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global hpcurrent As Integer                 'used during a battle, if this value drops to zero, the player has died'
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'question stats                                                                     ''''''''''''''''''''''''''''''''''''
    Global totalq, currentquestion, correctall, wrongall, correcttemp, wrongtemp As Integer 'Keeps status of current quiz(s)   '
                                                                                    ''''''''''''''''''''''''''''''''''''
'proggress flags (for save game option)         ''''''''''''''''''''''''''''''''''''
    Global currentlevel As Integer              'Variable to hold game proggression'
                                                ''''''''''''''''''''''''''''''''''''
'database Linking Variables                                 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Global searchstring, searchwhat, finalsearch As String  'Very important variable to search database according to randomised pattern'
                                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
    Public Type tStar           '''''''''''''''''''''''''''''''''''''''''''''''''
        nX As Long              'User defined variable to record star positions '
        nY As Long              'Saves each stars last location and speed so    '
        nSpeed As Integer       'it can be modified easily rather than          '
        nColor As Long          're-calculating EACH STAR EVERY LOOP            '
    End Type                    '''''''''''''''''''''''''''''''''''''''''''''''''

    Public G_dStar(1000) As tStar

Public Sub randomisequestions()

RandomisationProcess

nextquestion

End Sub

Sub RandomisationProcess()

'A simple loop to generate an array of 10 random numbers without duplicates
'...i know i should replece

    loopcounter = 0
    firstloopflag = 0
    Randomize
    Do
        b = Int((10 - 1 + 1) * Rnd + 1)
        If firstloopflag <> 0 Then
        loopcounter = loopcounter + 1
        End If
            If a(0) = b Or a(1) = b Or a(2) = b Or a(3) = b Or a(4) = b Or a(5) = b Or a(6) = b Or a(7) = b Or a(8) = b Or a(9) = b Then
                loopcounter = loopcounter - 1
            Else                                    '''''''''''''''''''''''''''''''''''''''''
                a(loopcounter) = b                  'Generates 10 unique random numbers and '
            End If                                  'saves them in he array a(), i might    '
        firstloopflag = 1                           'replace the if with a loop later on.    '
    Loop Until loopcounter > 8                      '''''''''''''''''''''''''''''''''''''''''
    
    loopcounter = 0
    
End Sub
Public Sub nextquestion()

'A simple bit of code to search the database according to the pre-roled random numbers
'...thus displaying the questions in a random order, the error capture just handles 11 in the array
'...because it is increased before this process

On Error GoTo 20:
    frmQuestions.dtaEasyLink.Recordset.MoveFirst
    searchstring = a(loopcounter)
    searchwhat = "[Question Number] = '"
    finalsearch = searchwhat & searchstring & "'"
    frmQuestions.dtaEasyLink.Recordset.FindNext (finalsearch)
    Exit Sub
20:

End Sub

Sub MeassageQuit()
'Quit message box
    frmMessageBox.lblMessage.Caption = "Are you sure you whant to quit?"
    frmMessageBox.optoption1.Caption = "Yes"
    frmMessageBox.optoption1.Visible = True
    frmMessageBox.optOption2.Visible = True
    frmMessageBox.optOption2.Caption = "No"
    frmMessageBox.Show vbModal
    
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then End
End Sub

Sub MeassageEnterLocation()
    frmMessageBox.lblMessage.Caption = "" & EnterMessageText & ""
    frmMessageBox.optoption1.Caption = "Yes"
    frmMessageBox.optoption1.Visible = True
    frmMessageBox.optOption2.Visible = True
    frmMessageBox.optOption2.Caption = "No"
    frmMessageBox.Show vbModal
    If EnterMessageText = "Do you wish to challenge me? 10 question according to your inteligence." Then
        FrmMain.setquestionsforgo
    Else
        If MBoxReturn = False Then
            frmWorldMap.picShip.Left = frmWorldMap.picShip.Left + 160
            Unload frmMessageBox
        End If
        If MBoxReturn = True Then initialisequestions
        EnterMessageText = ""
    End If
End Sub

Sub OKMessage()
'Cutom message box settings to simulate a vbOKonly msgbox
        frmMessageBox.lblMessage.Caption = "" & EnterMessageText & ""
        frmMessageBox.optOption2.Caption = "Ok"
        frmMessageBox.optoption1.Visible = False
        frmMessageBox.Show vbModal
        If MBoxReturn = False Then Exit Sub
        If MBoxReturn = True Then End
        EnterMessageText = ""
End Sub

Sub gameover()
    Unload frmbATTLE
    Unload FrmMain
    Unload frmmakenew
    frmGameOver.Show
    frmWorldMap.Visible = False
    WholeProcessTrigger = False
End Sub
Sub Yes_No()
'Cutom message box settings to simulate a vbYesNo msgbox

        frmMessageBox.lblMessage.Caption = "" & EnterMessageText & ""
        frmMessageBox.optOption2.Caption = "No"
        frmMessageBox.optoption1.Caption = "Yes"
        frmMessageBox.optOption2.Visible = True
        frmMessageBox.optoption1.Visible = True
        frmMessageBox.Show vbModal
        EnterMessageText = ""
End Sub
