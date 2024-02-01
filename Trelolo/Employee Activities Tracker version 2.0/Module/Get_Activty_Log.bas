Attribute VB_Name = "Get_Activty_Log"
Option Explicit

Public TimerActive As Boolean
Public FirstLoginTime As Date
Public CurrentActivityTime As Date
Public PreviousLoginHours As Date


'To update Activity and Other hours on all the cards available on Activity Tracker

Sub Timer() ' Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
   
    If TimerActive Then
        
        Dim TotalTime, ActivityTime
        
        TotalTime = Format(((Now() - FirstLoginTime) + PreviousLoginHours), "HH:MM:SS")
        
        If CurrentActivityTime > 0 Then
            
          ActivityTime = Format(((Now() - CurrentActivityTime)), "HH:MM:SS")
          ActivityTracker.CurrentActivityHours.Caption = ActivityTime
            
        End If
        
        Application.OnTime Now() + TimeValue("00:00:01"), "Timer"
        ActivityTracker.WorkingHours.Caption = TotalTime
        
        
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
End Sub


'Get Activity Log and Linking the raw source to List Box available on Activity Tracker Form

Sub Get_ActivityLog() ' Updated for Version 2.0
    
    Application.EnableCancelKey = xlDisabled
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
       
    Dim i As Integer
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Activity_Log")
    
    sh.Cells.Delete
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String
     
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
     
    qry = "Select * From tblActivityLog WHERE Dates=#" & VBA.Format(Date, "DD-MMMM-YYYY") & "#"
    qry = qry & " AND [Employee ID]='" & VBA.UCase(['Login Details'!$A$2]) & "'"
    qry = qry & " ORDER BY [Submitted On] DESC"

    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("a2").CopyFromRecordset rst
          
    For i = 1 To rst.Fields.Count
        sh.Cells(1, i).Value = VBA.UCase(rst.Fields(i - 1).Name)
    Next i
        
    rst.Close
    cnn.Close
      
    'Applying formatting to columns
    sh.Range("B:B").NumberFormat = "D-MMM-YY"
    sh.Range("J:K").NumberFormat = "D-MMM-YY HH:MM:SS AM/PM"
    sh.Range("L:L").NumberFormat = "HH:MM:SS"
    sh.Range("N:N").NumberFormat = "D-MMM-YY HH:MM:SS AM/PM"
    
         
    'Intializing List Box control
         
    With ActivityTracker.lstActivity
       
        .ColumnCount = 14
        .ColumnHeads = True
        .ColumnWidths = "0,50,90,90,0,0,0,100,0,75,75,75,0,0"
        
    
    End With
    
    
    PreviousLoginHours = [SUMIF(Activity_Log!$H:$H, "Login",Activity_Log!$L:$L)] ' Calculating Previous Login Hours
    
    ActivityTracker.WorkingHours.Caption = Format(PreviousLoginHours, "HH:MM:SS")
        
    ActivityTracker.ActivityHours.Caption = Format([SUMIF(Activity_Log!$H:$H, "<>Login",Activity_Log!$L:$L)] - [SUMIF(Activity_Log!$H:$H, "*Break*",Activity_Log!$L:$L)], "HH:MM:SS")
    
    ActivityTracker.AvailTime.Caption = _
    Format((PreviousLoginHours - [SUMIF(Activity_Log!$H:$H, "<>Login",Activity_Log!$L:$L)]), "HH:MM:SS")
          
    'Assigning RowSource to Listbox control
    If sh.Range("A" & Application.Rows.Count).End(xlUp).Row > 1 Then
        ActivityTracker.lstActivity.RowSource = "Activity_Log!A2:N" & sh.Range("A" & Application.Rows.Count).End(xlUp).Row
    Else
        ActivityTracker.lstActivity.RowSource = "Activity_Log!A2:N2"
    End If
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True


End Sub
