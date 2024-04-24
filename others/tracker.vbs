Dim sh_cha As Worksheet
Dim startTime As Date
Dim pausedTime As Date
Dim endTime As Date
Dim elapsed As Date
Dim timerRunning As Boolean
Dim stopTimerClicked As Boolean

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If Target.Range.Address = ThisWorkbook.Sheets("Sheet1").Range("I1").Address Then
        If timerRunning And Not stopTimerClicked Then
            ' Stop timer
            endTime = Now()
            elapsed = endTime - startTime
            ThisWorkbook.Sheets("Sheet1").Range("G1").Value = Format(endTime, "hh:mm:ss")
            ThisWorkbook.Sheets("Sheet1").Range("H1").Value = Format(elapsed, "hh:mm:ss")
            ThisWorkbook.Sheets("Sheet1").Hyperlinks.Add Anchor:=ThisWorkbook.Sheets("Sheet1").Range("I1"), Address:="", SubAddress:="I1", TextToDisplay:="Start Timer"
            timerRunning = False
            stopTimerClicked = True
        End If
    End If
End Sub

Sub Start_Time_cha()
    Set sh_cha = ThisWorkbook.Sheets("Sheet1")
    
    If Not timerRunning Then
        Dim startCell As Range
        Set startCell = sh_cha.Cells(sh_cha.Rows.Count, "F").End(xlUp) ' Find the last start time
        
        If startCell.Row > 1 Then
            ' Check if the cell above contains a valid start time
            Dim aboveCell As Range
            Set aboveCell = startCell.Offset(-1, 0)
            If IsDate(aboveCell.Value) Then
                startTime = aboveCell.Value
            Else
                startTime = Now()
            End If
        Else
            startTime = Now()
        End If
        
        ' Start timer
        startCell.Offset(1, 0).Value = Format(startTime, "hh:mm:ss")
        sh_cha.Hyperlinks.Add Anchor:=sh_cha.Cells(startCell.Row + 1, "I"), Address:="", SubAddress:="I1", TextToDisplay:="Stop Timer"
        timerRunning = True
        stopTimerClicked = False
        ' Update the timer display every second
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
    Else
        ' Pause the current timer
        PauseTimer
    End If
End Sub

Sub PauseTimer()
    If timerRunning Then
        ' Pause the timer
        pausedTime = Now()
        timerRunning = False
    End If
End Sub

Sub StopTimer()
    If timerRunning And Not stopTimerClicked Then
        ' Find the first empty cell in column G
        Dim lastRowG As Long
        lastRowG = sh_cha.Cells(sh_cha.Rows.Count, "G").End(xlUp).Row
        If lastRowG < 1 Then lastRowG = 1 ' In case there are no previous end times
        lastRowG = lastRowG + 1 ' Move to the next row
        
        ' Stop timer
        endTime = Now()
        elapsed = endTime - startTime
        sh_cha.Cells(lastRowG, "G").Value = Format(endTime, "hh:mm:ss")
        sh_cha.Cells(lastRowG, "H").Value = Format(elapsed, "hh:mm:ss")
        ThisWorkbook.Sheets("Sheet1").Hyperlinks.Add Anchor:=ThisWorkbook.Sheets("Sheet1").Range("I1"), Address:="", SubAddress:="A1", TextToDisplay:="Start Timer"
        timerRunning = False
        stopTimerClicked = True
    End If
End Sub

Sub UpdateTimer()
    If timerRunning Then
        ' Find the first empty cell in column F
        lastRowF = sh_cha.Cells(sh_cha.Rows.Count, "F").End(xlUp).Row
        If lastRowF < 1 Then lastRowF = 1 ' In case there are no previous start times
        
        elapsed = Now() - startTime
        sh_cha.Cells(lastRowF, "F").Value = Format(startTime, "hh:mm:ss") & " - " & Format(Now(), "hh:mm:ss")
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
    ElseIf pausedTime > 0 Then
        ' Update the paused time
        Dim pausedDuration As Date
        pausedDuration = pausedTime - startTime
        sh_cha.Cells(lastRowF, "F").Value = Format(startTime, "hh:mm:ss") & " - " & Format(pausedTime, "hh:mm:ss") & " (Paused: " & Format(pausedDuration, "hh:mm:ss") & ")"
        pausedTime = 0
        timerRunning = True
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
    End If
End Sub


