Attribute VB_Name = "CheckMapLocations"
Dim purey, purex As Integer
Sub locationcheck()

'Draws the cross hairs and cricles around where they cross

frmWorldMap.shpHorizontal.x1 = 0
frmWorldMap.shpHorizontal.X2 = frmWorldMap.ScaleWidth
frmWorldMap.shpHorizontal.y1 = frmWorldMap.picShip.Top + frmWorldMap.picShip.Height / 2
frmWorldMap.shpHorizontal.Y2 = frmWorldMap.picShip.Top + frmWorldMap.picShip.Height / 2
purex = frmWorldMap.shpHorizontal.y1

frmWorldMap.shpvert.y1 = 0
frmWorldMap.shpvert.Y2 = frmWorldMap.ScaleHeight
frmWorldMap.shpvert.X2 = frmWorldMap.picShip.Left + frmWorldMap.picShip.Width / 2
frmWorldMap.shpvert.x1 = frmWorldMap.picShip.Left + frmWorldMap.picShip.Width / 2
purey = frmWorldMap.shpvert.x1

frmWorldMap.shpCircle.Top = purex - frmWorldMap.shpCircle.Height / 2
frmWorldMap.shpCircle.Left = purey - frmWorldMap.shpCircle.Width / 2
frmWorldMap.shpcompass.x1 = purey
frmWorldMap.shpcompass.y1 = purex

'draws the little white line, in the direction that thrust will move
frmWorldMap.shpcompass.X2 = intX1
frmWorldMap.shpcompass.Y2 = intY1

Unload frmMessageBox

'frmWorldmap.level 1

If frmWorldMap.picShip.Left > frmWorldMap.Level1.Left And frmWorldMap.picShip.Left < frmWorldMap.Level1.Left + frmWorldMap.Level1.Width And frmWorldMap.picShip.Top > frmWorldMap.Level1.Top And frmWorldMap.picShip.Top < frmWorldMap.Level1.Top + frmWorldMap.Level1.Height Then
    question1
    GoTo 20
End If

'frmWorldmap.level 2

If frmWorldMap.picShip.Left > frmWorldMap.Level2.Left And frmWorldMap.picShip.Left < frmWorldMap.Level2.Left + frmWorldMap.Level2.Width And frmWorldMap.picShip.Top > frmWorldMap.Level2.Top And frmWorldMap.picShip.Top < frmWorldMap.Level2.Top + frmWorldMap.Level2.Height Then
    question2
    GoTo 20
End If


'frmWorldmap.level 3

If frmWorldMap.picShip.Left > frmWorldMap.Level3.Left And frmWorldMap.picShip.Left < frmWorldMap.Level3.Left + frmWorldMap.Level3.Width And frmWorldMap.picShip.Top > frmWorldMap.Level3.Top And frmWorldMap.picShip.Top < frmWorldMap.Level3.Top + frmWorldMap.Level3.Height Then
    question3
    GoTo 20
End If


'frmWorldmap.level 4

If frmWorldMap.picShip.Left > frmWorldMap.Level4.Left And frmWorldMap.picShip.Left < frmWorldMap.Level4.Left + frmWorldMap.Level4.Width And frmWorldMap.picShip.Top > frmWorldMap.Level4.Top And frmWorldMap.picShip.Top < frmWorldMap.Level4.Top + frmWorldMap.Level4.Height Then
    question4
    GoTo 20
End If


'frmWorldmap.level 5

If frmWorldMap.picShip.Left > frmWorldMap.Level5.Left And frmWorldMap.picShip.Left < frmWorldMap.Level5.Left + frmWorldMap.Level5.Width And frmWorldMap.picShip.Top > frmWorldMap.Level5.Top And frmWorldMap.picShip.Top < frmWorldMap.Level5.Top + frmWorldMap.Level5.Height Then
    question5
    GoTo 20
End If


'frmWorldmap.level 6

If frmWorldMap.picShip.Left > frmWorldMap.Level6.Left And frmWorldMap.picShip.Left < frmWorldMap.Level6.Left + frmWorldMap.Level6.Width And frmWorldMap.picShip.Top > frmWorldMap.Level6.Top And frmWorldMap.picShip.Top < frmWorldMap.Level6.Top + frmWorldMap.Level6.Height Then
    question6
    GoTo 20
End If


'frmWorldmap.level 7

If frmWorldMap.picShip.Left > frmWorldMap.Level7.Left And frmWorldMap.picShip.Left < frmWorldMap.Level7.Left + frmWorldMap.Level7.Width And frmWorldMap.picShip.Top > frmWorldMap.Level7.Top And frmWorldMap.picShip.Top < frmWorldMap.Level7.Top + frmWorldMap.Level7.Height Then
    question7
    GoTo 20
End If


'frmWorldmap.level 8

If frmWorldMap.picShip.Left > frmWorldMap.Level8.Left And frmWorldMap.picShip.Left < frmWorldMap.Level8.Left + frmWorldMap.Level8.Width And frmWorldMap.picShip.Top > frmWorldMap.Level8.Top And frmWorldMap.picShip.Top < frmWorldMap.Level8.Top + frmWorldMap.Level8.Height Then
    question8
    GoTo 20
End If


'frmWorldmap.level 9

If frmWorldMap.picShip.Left > frmWorldMap.Level9.Left And frmWorldMap.picShip.Left < frmWorldMap.Level9.Left + frmWorldMap.Level9.Width And frmWorldMap.picShip.Top > frmWorldMap.Level9.Top And frmWorldMap.picShip.Top < frmWorldMap.Level9.Top + frmWorldMap.Level9.Height Then
    question9
    GoTo 20
End If


'frmWorldmap.level 10

If frmWorldMap.picShip.Left > frmWorldMap.Level10.Left And frmWorldMap.picShip.Left < frmWorldMap.Level10.Left + frmWorldMap.Level10.Width And frmWorldMap.picShip.Top > frmWorldMap.Level10.Top And frmWorldMap.picShip.Top < frmWorldMap.Level10.Top + frmWorldMap.Level10.Height Then
    question10
    GoTo 20
End If


'Home

If frmWorldMap.picShip.Left > frmWorldMap.home.Left And frmWorldMap.picShip.Left < frmWorldMap.home.Left + frmWorldMap.home.Width And frmWorldMap.picShip.Top > frmWorldMap.home.Top And frmWorldMap.picShip.Top < frmWorldMap.home.Top + frmWorldMap.home.Height Then
    question11
    GoTo 20
End If
20:
End Sub

Sub question1()
    EnterMessageText = "Mission 1: Thengal Temple, Do You Wish To Enter The Thengal Province?"
    currentlevel = 1
    MeassageEnterLocation
End Sub
Sub question2()
    EnterMessageText = "Mission 2: Xu Bay, Do You Wish To Enter and challenge the question master in the subject of Biology?"
    currentlevel = 2
    MeassageEnterLocation
End Sub
Sub question3()
    EnterMessageText = "Mission 3: Ten Prison, Do You Wish To Enter and challenge the question master in the subject of Mathematics?"
    currentlevel = 3
    MeassageEnterLocation
End Sub
Sub question4()
    EnterMessageText = "Mission 4: Sh-Hi-Na, Do You Wish To Enter and challenge the question master in the subject of Physics?"
    currentlevel = 4
    MeassageEnterLocation
End Sub
Sub question5()
    EnterMessageText = "Mission 5: MoRR, Do You Wish To Enter and challenge the question master in the subject of Nature?"
    currentlevel = 5
    MeassageEnterLocation
End Sub
Sub question6()
    EnterMessageText = "Mission 6: Tronx Deep-Sea Research Center, Do You Wish To Enter and challenge the question master in the subject of Geography?"
    currentlevel = 6
    MeassageEnterLocation
End Sub
Sub question7()
    EnterMessageText = "Mission 7: Es-AssA, Do You Wish To Enter and challenge the question master in the subject of Riddles?"
    currentlevel = 7
    MeassageEnterLocation
End Sub
Sub question8()
    EnterMessageText = "Mission 8: Shelandor Province, Do You Wish To Enter and challenge the question master in the subject of Mechanics?"
    currentlevel = 8
    MeassageEnterLocation
End Sub
Sub question9()
    EnterMessageText = "Mission 9: Zanori Tessa, Do You Wish To Enter and challenge the question master in the subject of Media?"
    currentlevel = 9
    MeassageEnterLocation
End Sub
Sub question10()
    EnterMessageText = "Mission 10: The Lost Library, Do You Wish To Enter and challenge the question master in the subject of Music?"
    currentlevel = 10
    MeassageEnterLocation
End Sub

Sub question11()
    reply = MsgBox("Home Village: Do You Wish To Enter?", vbYesNo, "The Quest Of Knowledge")
    If reply = 6 Then
        
    Else
        frmWorldMap.picShip.Left = frmWorldMap.picShip.Left + 160
    End If
End Sub
Sub initialisequestions()

        FrmMain.Show
        FrmMain.imgPortrait.Picture = frmWorldMap.imgPortrait.Picture

End Sub
