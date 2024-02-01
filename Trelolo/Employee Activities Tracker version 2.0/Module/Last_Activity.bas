Attribute VB_Name = "Last_Activity"

Option Explicit

'Code to load last activity or login if tool has not been properly closed or logged out

Sub Load_Incompleted_Log() ' Updated Version 2.0
    
    Dim iRow As Double

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler

    Dim sh As Worksheet
    
    Set sh = ThisWorkbook.Sheets("Activity_Log")

    
    If Trim(sh.Range("J2").Value) <> "" And Trim(sh.Range("K2").Value) = "" Then
    
        MsgBox "There was an unexpected logout. Loading previous pending activity.", vbOKOnly + vbCritical, "Unexpected Logout"
        
        'Fetching Last Login
        Call ActivityTracker.Enable_Activity_Control
        ActivityTracker.cmdLogin.Enabled = False
        ActivityTracker.cmdLogout.Enabled = True
        FirstLoginTime = [VLOOKUP("Login",Activity_Log!$H:$K,3,0)]
        ActivityTracker.LoginTime.Caption = Format(FirstLoginTime, "HH:MM:SS AM/PM")
        TimerActive = True 'Start time
        Call Timer
        
        'Updating Details
         
         ActivityTracker.txtDescription.Value = sh.Range("I2").Value             'ACTIVITY DESCRIPTION
         
         If VBA.UCase(Trim(sh.Range("H2").Value)) <> VBA.UCase("Login") Then
            CurrentActivityTime = sh.Range("J2").Value                              'StartTime
            
            ActivityTracker.cmbClientName.Value = sh.Range("F2").Value              'Client Name
            ActivityTracker.cmbLocationName.Value = sh.Range("G2").Value            'Location
            ActivityTracker.lstActivityCode.Value = sh.Range("H2").Value            'ACTIVITY TYPE
            Call Lock_UserInput
         End If
            
    ElseIf [IFERROR(VLOOKUP("Login",Activity_Log!$H:$K,4,0),"NA")] = 0 Then         ' if Logout time not captured
    
        iRow = [Match("Login",Activity_Log!$H:$H,0)]
        MsgBox "There was an unexpected logout. Loading previous pending activity.", vbOKOnly + vbCritical, "Unexpected Logout"
        
        'Fetching Last Login
        Call ActivityTracker.Enable_Activity_Control
        ActivityTracker.cmdLogin.Enabled = False
        ActivityTracker.cmdLogout.Enabled = True
        FirstLoginTime = [VLOOKUP("Login",Activity_Log!$H:$K,3,0)]
        ActivityTracker.LoginTime.Caption = Format(FirstLoginTime, "HH:MM:SS AM/PM")
        TimerActive = True                                                             'Start time
        Call Timer
        
        'Updating Details
        
        ActivityTracker.txtDescription.Value = sh.Range("I2").Value                     'ACTIVITY DESCRIPTION
        
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
          
End Sub
