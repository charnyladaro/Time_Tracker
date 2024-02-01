Attribute VB_Name = "Update_Activity"
Option Explicit

'Code to Add or Update the Login, Activity, Break and Logout  to Database

Sub Add_StartEntry(JobCode As String, TimeType As Boolean, TimeStamp As Variant) 'Updated for Version 2.0
    
    Application.EnableCancelKey = xlDisabled
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim qry As String
    Dim sqry As String
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
        
    ' TimeType -      True - StartTime; False- EndTime
    
    If TimeType = True Then
    
        qry = "Select * from tblActivityLog"
        
    Else
    
        'query
        qry = "Select * From tblActivityLog WHERE Dates=#" & VBA.Format(Date, "DD-MMMM-YYYY") & "#"
        qry = qry & " AND [Employee ID]='" & VBA.UCase(['Login Details'!$A$2]) & "'"
        qry = qry & " AND [ACTIVITY TYPE]='" & JobCode & "'"
        qry = qry & " AND [END TIME] is NULL"
        'qry = qry & " AND [Serial_No] IN (" & sqry & ")"
        qry = qry & " ORDER BY [Serial_No]"
        
    End If
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    If rst.RecordCount > 0 Then
        rst.MoveLast
    End If
    
    With rst
       
       If TimeType = True Then
       
         .AddNew ' If new records
       
        .Fields("DATES").Value = ActivityTracker.txtDate.Value
        .Fields("Employee ID").Value = VBA.UCase(['Login Details'!$A$2])
        .Fields("EMPLOYEE NAME").Value = ActivityTracker.txtEmployeeName
        .Fields("Supervisor Name").Value = ActivityTracker.txtSupervisor
        
        .Fields("Client Name").Value = ActivityTracker.cmbClientName
        .Fields("Location").Value = ActivityTracker.cmbLocationName
        .Fields("ACTIVITY TYPE").Value = JobCode
        .Fields("ACTIVITY DESCRIPTION").Value = ActivityTracker.txtDescription
        
        .Fields("START TIME").Value = TimeStamp
        .Fields("SUBMITTED BY").Value = Environ("username")
        .Fields("SUBMITTED ON").Value = [Now()]
        
        
        .Update
       
       Else
        
        If rst.RecordCount > 0 Then
        
        .Fields("END TIME").Value = TimeStamp
        .Fields("TOTAL TIME").Value = .Fields("END TIME").Value - .Fields("START TIME").Value
        .Fields("SUBMITTED BY").Value = Environ("username")
        .Fields("SUBMITTED ON").Value = [Now()]
        
        
        .Update
        
        Else
        
            MsgBox "No corresponding records found! Please retry again.", vbOKOnly, "Records not found"
        
        End If
        
       End If
            
       .Close
    
    End With
    
    cnn.Close
    
    'Call Procedure to Updated the List Box with Newly added dataset
    Call Get_ActivityLog
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub


