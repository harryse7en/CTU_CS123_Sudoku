Imports System.IO


Public Class MainForm

    'Where the saved game will be going to
    Dim FileSaveName As String = "SudokuGameData.txt"
    Dim PlayTime As Integer



    Private Sub about_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutBtn.Click
        'Linked to about button
        MsgBox("Created by harryse7en" & (Chr(13)) & "for CTU class CS123" & (Chr(13)) & "2/2/2008")
    End Sub



    Private Sub btnup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles B1.Click, B2.Click, B3.Click, B4.Click, B5.Click, B6.Click, B7.Click, B8.Click, B9.Click, B10.Click, B11.Click, B12.Click, B13.Click, B14.Click, B15.Click, B16.Click
        'Starts play timer
        PlayTimer.Enabled = True
        TextBox2.Text = "Can you solve me?"
        'Increments each button when clicked
        If sender.text = "" Then
            sender.text = 1
            Exit Sub
        ElseIf sender.text = 1 Then
            sender.text = 2
            Exit Sub
        ElseIf sender.text = 2 Then
            sender.text = 3
            Exit Sub
        ElseIf sender.text = 3 Then
            sender.text = 4
            Exit Sub
        ElseIf sender.text = 4 Then
            sender.text = ""
            Exit Sub
        ElseIf sender.text > 4 Then
            sender.text = ""
            Exit Sub
        End If
    End Sub



    Private Sub check_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBtn.Click
        'Starts play timer, if user clicks Check It instead of game board
        PlayTimer.Enabled = True
        'Check each row
        If B1.Text = B2.Text Or B1.Text = B3.Text Or B1.Text = B4.Text Or B2.Text = B3.Text Or B2.Text = B4.Text Or B3.Text = B4.Text Then
            TextBox2.Text = "There is a problem in row 1!"
            Exit Sub
        ElseIf B5.Text = B6.Text Or B5.Text = B7.Text Or B5.Text = B8.Text Or B6.Text = B7.Text Or B6.Text = B8.Text Or B7.Text = B8.Text Then
            TextBox2.Text = "There is a problem in row 2!"
            Exit Sub
        ElseIf B9.Text = B10.Text Or B9.Text = B11.Text Or B9.Text = B12.Text Or B10.Text = B11.Text Or B10.Text = B12.Text Or B11.Text = B12.Text Then
            TextBox2.Text = "There is a problem in row 3!"
            Exit Sub
        ElseIf B13.Text = B14.Text Or B13.Text = B15.Text Or B13.Text = B16.Text Or B14.Text = B15.Text Or B14.Text = B16.Text Or B15.Text = B16.Text Then
            TextBox2.Text = "There is a problem in row 4!"
            Exit Sub
            'Check each column
        ElseIf B1.Text = B5.Text Or B1.Text = B9.Text Or B1.Text = B13.Text Or B5.Text = B9.Text Or B5.Text = B13.Text Or B9.Text = B13.Text Then
            TextBox2.Text = "There is a problem in column 1!"
            Exit Sub
        ElseIf B2.Text = B6.Text Or B2.Text = B10.Text Or B2.Text = B14.Text Or B6.Text = B10.Text Or B6.Text = B14.Text Or B10.Text = B14.Text Then
            TextBox2.Text = "There is a problem in column 2!"
            Exit Sub
        ElseIf B3.Text = B7.Text Or B3.Text = B11.Text Or B3.Text = B15.Text Or B7.Text = B11.Text Or B7.Text = B15.Text Or B11.Text = B15.Text Then
            TextBox2.Text = "There is a problem in column 3!"
            Exit Sub
        ElseIf B4.Text = B8.Text Or B4.Text = B12.Text Or B4.Text = B16.Text Or B8.Text = B12.Text Or B8.Text = B16.Text Or B12.Text = B16.Text Then
            TextBox2.Text = "There is a problem in column 4!"
            Exit Sub
            'Check each quadrant
        ElseIf B1.Text = B2.Text Or B1.Text = B5.Text Or B1.Text = B6.Text Or B2.Text = B5.Text Or B2.Text = B6.Text Or B5.Text = B6.Text Then
            TextBox2.Text = "There is a problem in quadrant 1!"
            Exit Sub
        ElseIf B3.Text = B4.Text Or B3.Text = B7.Text Or B3.Text = B8.Text Or B4.Text = B7.Text Or B4.Text = B8.Text Or B7.Text = B8.Text Then
            TextBox2.Text = "There is a problem in quadrant 2!"
            Exit Sub
        ElseIf B9.Text = B10.Text Or B9.Text = B13.Text Or B9.Text = B14.Text Or B10.Text = B13.Text Or B10.Text = B14.Text Or B13.Text = B14.Text Then
            TextBox2.Text = "There is a problem in quadrant 3!"
            Exit Sub
        ElseIf B11.Text = B12.Text Or B11.Text = B15.Text Or B11.Text = B16.Text Or B12.Text = B15.Text Or B12.Text = B16.Text Or B15.Text = B16.Text Then
            TextBox2.Text = "There is a problem in quadrant 4!"
            Exit Sub
            'Check for blanks
        ElseIf B1.Text = "" Or B2.Text = "" Or B3.Text = "" Or B4.Text = "" Or B5.Text = "" Or B6.Text = "" Or B7.Text = "" Or B8.Text = "" Or B9.Text = "" Or B10.Text = "" Or B11.Text = "" Or B12.Text = "" Or B13.Text = "" Or B14.Text = "" Or B15.Text = "" Or B16.Text = "" Then
            TextBox2.Text = "Good so far, but there are blanks!"
            'Present won message
        Else
            PlayTimer.Enabled = False
            TextBox2.Text = "YOU WIN!!!"
            MsgBox("Congrats to you, for you have won the game!" & (Chr(13)) & "It took you " & PlayTime & " seconds.")
            'Resets playtimer
            PlayTime = 0
            TextBox1.Text = ""
            Exit Sub
        End If
    End Sub



    Private Sub PlayTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlayTimer.Tick
        'Increments playtime every 1000 ms and displays it
        PlayTime = PlayTime + 1
        TextBox1.Text = PlayTime
    End Sub



    Private Sub quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuitBtn.Click
        'Linked to exit button
        Dim Confirm As MsgBoxResult
        Confirm = MsgBox("Do you really want to quit?", MsgBoxStyle.YesNo)
        If Confirm = MsgBoxResult.Yes Then
            Me.Close()
        Else : Exit Sub
        End If
    End Sub



    Private Sub newgame()
        'Linked to random button, creates a random board for the user
        Dim RB(15) As Integer
        Dim RandBlanks(10) As Integer
        Dim Finished As Integer
        Dim Done As Integer
        Dim i As Integer
        Dim j As Integer

        'Assigns 1 to each game box
        Do Until i > 15
            RB(i) = 1
            i = i + 1
        Loop

        'Assign variables to 0
        i = 0
        j = 0
        Finished = 0

        'Randomizer
        GoTo Startup
Startup:
        Do While Finished = 0
            Done = 0
            GoTo rand1
rand1:
            'Random row one
            Do Until RB(0) <> RB(1) And RB(0) <> RB(2) And RB(0) <> RB(3) And RB(1) <> RB(2) And RB(1) <> RB(3) And RB(2) <> RB(3) And Done = 1
                RB(0) = CInt(Int(Rnd() * 4 + 1))
                RB(1) = CInt(Int(Rnd() * 4 + 1))
                RB(2) = CInt(Int(Rnd() * 4 + 1))
                RB(3) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 100 Then
                    i = 0
                    GoTo rand1
                End If
            Loop
            Done = 0
            i = 0
            GoTo rand2
rand2:
            'Random row two
            Do Until RB(4) <> RB(5) And RB(4) <> RB(6) And RB(4) <> RB(7) And RB(5) <> RB(6) And RB(5) <> RB(7) And RB(6) <> RB(7) And RB(0) <> RB(4) And RB(1) <> RB(5) And RB(2) <> RB(6) And RB(3) <> RB(7) And RB(0) <> RB(5) And RB(2) <> RB(7) And Done = 1
                RB(4) = CInt(Int(Rnd() * 4 + 1))
                RB(5) = CInt(Int(Rnd() * 4 + 1))
                RB(6) = CInt(Int(Rnd() * 4 + 1))
                RB(7) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 100 Then
                    i = 0
                    GoTo rand2
                End If
            Loop
            Done = 0
            i = 0
            GoTo rand3
rand3:
            Do Until RB(0) <> RB(8) And RB(4) <> RB(8) And Done = 1
                RB(8) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 100 Then
                    i = 0
                    GoTo rand2
                End If
            Loop
            Done = 0
            i = 0
            GoTo rand3b
rand3b:
            Do Until RB(8) <> RB(9) And RB(1) <> RB(9) And RB(5) <> RB(9) And Done = 1
                RB(9) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 200 Then
                    i = 0
                    GoTo rand3
                End If
            Loop
            Done = 0
            i = 0
            GoTo rand4
rand4:
            Do Until RB(10) <> RB(8) And RB(10) <> RB(9) And RB(2) <> RB(10) And RB(6) <> RB(10) And Done = 1
                RB(10) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 200 Then
                    i = 0
                    GoTo rand3
                End If
            Loop
            Done = 0
            i = 0
            GoTo rand5
rand5:
            Do Until RB(11) <> RB(8) And RB(11) <> RB(9) And RB(10) <> RB(11) And RB(3) <> RB(11) And RB(7) <> RB(11) And Done = 1
                RB(11) = CInt(Int(Rnd() * 4 + 1))
                Done = 1
                i = i + 1
                If i > 200 Then
                    i = 0
                    GoTo rand3
                End If
            Loop
            GoTo lastrow
lastrow:
            'Assigns button 13 according to existing buttons
            If RB(0) <> 1 And RB(4) <> 1 And RB(8) <> 1 And RB(9) <> 1 Then
                RB(12) = 1
            ElseIf RB(0) <> 2 And RB(4) <> 2 And RB(8) <> 2 And RB(9) <> 2 Then
                RB(12) = 2
            ElseIf RB(0) <> 3 And RB(4) <> 3 And RB(8) <> 3 And RB(9) <> 3 Then
                RB(12) = 3
            ElseIf RB(0) <> 4 And RB(4) <> 4 And RB(8) <> 4 And RB(9) <> 4 Then
                RB(12) = 4
            Else
                Done = 0
                GoTo rand2
            End If
            'Assigns button 14 according to existing buttons
            If RB(1) <> 1 And RB(5) <> 1 And RB(9) <> 1 And RB(8) <> 1 And RB(12) <> 1 Then
                RB(13) = 1
            ElseIf RB(1) <> 2 And RB(5) <> 2 And RB(9) <> 2 And RB(8) <> 2 And RB(12) <> 1 Then
                RB(13) = 2
            ElseIf RB(1) <> 3 And RB(5) <> 3 And RB(9) <> 3 And RB(8) <> 3 And RB(12) <> 3 Then
                RB(13) = 3
            ElseIf RB(1) <> 4 And RB(5) <> 4 And RB(9) <> 4 And RB(8) <> 4 And RB(12) <> 4 Then
                RB(13) = 4
            Else
                Done = 0
                GoTo rand2
            End If
            'Assigns button 15 according to existing buttons
            If RB(2) <> 1 And RB(6) <> 1 And RB(10) <> 1 And RB(11) <> 1 And RB(12) <> 1 And RB(13) <> 1 Then
                RB(14) = 1
            ElseIf RB(2) <> 2 And RB(6) <> 2 And RB(10) <> 2 And RB(11) <> 2 And RB(12) <> 2 And RB(13) <> 2 Then
                RB(14) = 2
            ElseIf RB(2) <> 3 And RB(6) <> 3 And RB(10) <> 3 And RB(11) <> 3 And RB(12) <> 3 And RB(13) <> 3 Then
                RB(14) = 3
            ElseIf RB(2) <> 4 And RB(6) <> 4 And RB(10) <> 4 And RB(11) <> 4 And RB(12) <> 4 And RB(13) <> 4 Then
                RB(14) = 4
            Else
                Done = 0
                GoTo rand2
            End If
            'Assigns button 16 according to existing buttons
            If RB(3) <> 1 And RB(7) <> 1 And RB(11) <> 1 And RB(10) <> 1 And RB(12) <> 1 And RB(13) <> 1 And RB(14) <> 1 Then
                RB(15) = 1
            ElseIf RB(3) <> 2 And RB(7) <> 2 And RB(11) <> 2 And RB(10) <> 2 And RB(12) <> 2 And RB(13) <> 2 And RB(14) <> 2 Then
                RB(15) = 2
            ElseIf RB(3) <> 3 And RB(7) <> 3 And RB(11) <> 3 And RB(10) <> 3 And RB(12) <> 3 And RB(13) <> 3 And RB(14) <> 3 Then
                RB(15) = 3
            ElseIf RB(3) <> 4 And RB(7) <> 4 And RB(11) <> 4 And RB(10) <> 4 And RB(12) <> 4 And RB(13) <> 4 And RB(14) <> 4 Then
                RB(15) = 4
            Else
                Done = 0
                GoTo rand2
            End If
            'Counts do iterations, breaks an infinite loop after 100 cycles if it locks up
            j = j + 1
            If j > 100 Then
                GoTo Startup
            End If
            Finished = 1
        Loop

        'Publishes randomizer to game board
        B1.Text = RB(0)
        B2.Text = RB(1)
        B3.Text = RB(2)
        B4.Text = RB(3)
        B5.Text = RB(4)
        B6.Text = RB(5)
        B7.Text = RB(6)
        B8.Text = RB(7)
        B9.Text = RB(8)
        B10.Text = RB(9)
        B11.Text = RB(10)
        B12.Text = RB(11)
        B13.Text = RB(12)
        B14.Text = RB(13)
        B15.Text = RB(14)
        B16.Text = RB(15)

        'Chooses three random numbers for the three blanks
        Do Until j = 1 And RandBlanks(0) <> RandBlanks(1) And RandBlanks(0) <> RandBlanks(2) And RandBlanks(1) <> RandBlanks(2)
            RandBlanks(0) = CInt(Int(Rnd() * 16 + 1))
            RandBlanks(1) = CInt(Int(Rnd() * 16 + 1))
            RandBlanks(2) = CInt(Int(Rnd() * 16 + 1))
            j = 1
        Loop
        If RandBlanks(0) = 1 Or RandBlanks(1) = 1 Or RandBlanks(2) = 1 Then
            B1.Text = ""
        End If
        If RandBlanks(0) = 2 Or RandBlanks(1) = 2 Or RandBlanks(2) = 2 Then
            B2.Text = ""
        End If
        If RandBlanks(0) = 3 Or RandBlanks(1) = 3 Or RandBlanks(2) = 3 Then
            B3.Text = ""
        End If
        If RandBlanks(0) = 4 Or RandBlanks(1) = 4 Or RandBlanks(2) = 4 Then
            B4.Text = ""
        End If
        If RandBlanks(0) = 5 Or RandBlanks(1) = 5 Or RandBlanks(2) = 5 Then
            B5.Text = ""
        End If
        If RandBlanks(0) = 6 Or RandBlanks(1) = 6 Or RandBlanks(2) = 6 Then
            B6.Text = ""
        End If
        If RandBlanks(0) = 7 Or RandBlanks(1) = 7 Or RandBlanks(2) = 7 Then
            B7.Text = ""
        End If
        If RandBlanks(0) = 8 Or RandBlanks(1) = 8 Or RandBlanks(2) = 8 Then
            B8.Text = ""
        End If
        If RandBlanks(0) = 9 Or RandBlanks(1) = 9 Or RandBlanks(2) = 9 Then
            B9.Text = ""
        End If
        If RandBlanks(0) = 10 Or RandBlanks(1) = 10 Or RandBlanks(2) = 10 Then
            B10.Text = ""
        End If
        If RandBlanks(0) = 11 Or RandBlanks(1) = 11 Or RandBlanks(2) = 11 Then
            B11.Text = ""
        End If
        If RandBlanks(0) = 12 Or RandBlanks(1) = 12 Or RandBlanks(2) = 12 Then
            B12.Text = ""
        End If
        If RandBlanks(0) = 13 Or RandBlanks(1) = 13 Or RandBlanks(2) = 13 Then
            B13.Text = ""
        End If
        If RandBlanks(0) = 14 Or RandBlanks(1) = 14 Or RandBlanks(2) = 14 Then
            B14.Text = ""
        End If
        If RandBlanks(0) = 15 Or RandBlanks(1) = 15 Or RandBlanks(2) = 15 Then
            B15.Text = ""
        End If
        If RandBlanks(0) = 16 Or RandBlanks(1) = 16 Or RandBlanks(2) = 16 Then
            B16.Text = ""
        End If
        PlayTime = 0
        TextBox2.Text = "Can you solve me?"
    End Sub



    Private Sub save_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBtn.Click
        Dim SaveGame As New IO.StreamWriter(FileSaveName)
        SaveGame.WriteLine(B1.Text)
        SaveGame.WriteLine(B2.Text)
        SaveGame.WriteLine(B3.Text)
        SaveGame.WriteLine(B4.Text)
        SaveGame.WriteLine(B5.Text)
        SaveGame.WriteLine(B6.Text)
        SaveGame.WriteLine(B7.Text)
        SaveGame.WriteLine(B8.Text)
        SaveGame.WriteLine(B9.Text)
        SaveGame.WriteLine(B10.Text)
        SaveGame.WriteLine(B11.Text)
        SaveGame.WriteLine(B12.Text)
        SaveGame.WriteLine(B13.Text)
        SaveGame.WriteLine(B14.Text)
        SaveGame.WriteLine(B15.Text)
        SaveGame.WriteLine(B16.Text)
        SaveGame.Close()
        MsgBox("Your game was saved!")
    End Sub



    Private Sub load_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadBtn.Click
        Dim confirm As MsgBoxResult
        Dim LoadGame As New IO.StreamReader(FileSaveName)
        confirm = MsgBox("Do you want to load a saved game?" & (Chr(13)) & "Note: You will lose all progress on the current game!", MsgBoxStyle.OkCancel)
        If confirm = MsgBoxResult.Ok Then
            B1.Text = LoadGame.ReadLine()
            B2.Text = LoadGame.ReadLine()
            B3.Text = LoadGame.ReadLine()
            B4.Text = LoadGame.ReadLine()
            B5.Text = LoadGame.ReadLine()
            B6.Text = LoadGame.ReadLine()
            B7.Text = LoadGame.ReadLine()
            B8.Text = LoadGame.ReadLine()
            B9.Text = LoadGame.ReadLine()
            B10.Text = LoadGame.ReadLine()
            B11.Text = LoadGame.ReadLine()
            B12.Text = LoadGame.ReadLine()
            B13.Text = LoadGame.ReadLine()
            B14.Text = LoadGame.ReadLine()
            B15.Text = LoadGame.ReadLine()
            B16.Text = LoadGame.ReadLine()
            PlayTime = 0
            TextBox2.Text = "Can you solve me?"
            MsgBox("Game loaded successfully.")
            LoadGame.Close()
        Else : LoadGame.Close()
            Exit Sub
        End If
    End Sub



    Private Sub game1_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Game1.Click
        B1.Text = "1"
        B2.Text = "2"
        B3.Text = "3"
        B4.Text = "4"
        B5.Text = ""
        B6.Text = "4"
        B7.Text = "1"
        B8.Text = "2"
        B9.Text = ""
        B10.Text = "3"
        B11.Text = "4"
        B12.Text = "1"
        B13.Text = "4"
        B14.Text = "1"
        B15.Text = "2"
        B16.Text = ""
        PlayTime = 0
        TextBox2.Text = "Can you solve me?"
    End Sub



    Private Sub game2_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Game2.Click
        B1.Text = "1"
        B2.Text = "4"
        B3.Text = ""
        B4.Text = "2"
        B5.Text = "2"
        B6.Text = "3"
        B7.Text = "4"
        B8.Text = ""
        B9.Text = "4"
        B10.Text = "2"
        B11.Text = "1"
        B12.Text = "3"
        B13.Text = ""
        B14.Text = "1"
        B15.Text = "2"
        B16.Text = "4"
        PlayTime = 0
        TextBox2.Text = "Can you solve me?"
    End Sub



    Private Sub game3_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Game3.Click
        B1.Text = "4"
        B2.Text = "1"
        B3.Text = "3"
        B4.Text = "2"
        B5.Text = "3"
        B6.Text = "2"
        B7.Text = ""
        B8.Text = "4"
        B9.Text = "2"
        B10.Text = "3"
        B11.Text = "4"
        B12.Text = "1"
        B13.Text = ""
        B14.Text = "4"
        B15.Text = "2"
        B16.Text = ""
        PlayTime = 0
        TextBox2.Text = "Can you solve me?"
    End Sub



    Private Sub Random_click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBtn.Click
        newgame()
        TextBox2.Text = "Can you solve me?"
    End Sub
End Class
