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
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Your existing code here
    
    ' Check if the resume button shape exists
    If resumeButton Is Nothing Then
        ' Get the resume button shape
        For Each sh In sh_cha.Shapes
            If sh.TopLeftCell.Address = "$J$1" Then
                Set resumeButton = sh
                Exit For
            End If
        Next sh
    End If
    
    ' Hide the resume button initially
    If Not resumeButton Is Nothing Then
        resumeButton.Visible = False
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
        ' Pause the timer by saving the elapsed time so far
        pausedTime = Now()
        timerRunning = False
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
        
        ' Show the resume button
        resumeButton.Visible = True
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
        ' Timer is running, update the timer display
        lastRowF = sh_cha.Cells(sh_cha.Rows.Count, "F").End(xlUp).Row
        If lastRowF < 1 Then lastRowF = 1 ' In case there are no previous start times
        
        elapsed = Now() - startTime
        sh_cha.Cells(lastRowF, "F").Value = Format(startTime, "hh:mm:ss") & " - " & Format(Now(), "hh:mm:ss")
        Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
    ElseIf pausedTime > 0 Then
        ' Timer is paused, update the timer display to show the paused time
        lastRowF = sh_cha.Cells(sh_cha.Rows.Count, "F").End(xlUp).Row
        If lastRowF < 1 Then lastRowF = 1 ' In case there are no previous start times
        
        Dim pausedDuration As Date
        pausedDuration = pausedTime - startTime
        sh_cha.Cells(lastRowF, "F").Value = Format(startTime, "hh:mm:ss") & " - " & Format(pausedTime, "hh:mm:ss") & " (Paused: " & Format(pausedDuration, "hh:mm:ss") & ")"
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Check if the selection is the resume button
    If Not Intersect(Target, sh_cha.Range("J1")) Is Nothing Then
        ' Call the ResumeTimer subroutine
        ResumeTimer
    End If
End Sub

Sub ResumeTimer()
    ' Resume the timer
    Dim pausedDuration As Double
    pausedDuration = Now() - pausedTime
    startTime = startTime + pausedDuration
    pausedTime = 0
    timerRunning = True
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateTimer"
    
    ' Hide the resume button
    resumeButton.Visible = False
End Sub
