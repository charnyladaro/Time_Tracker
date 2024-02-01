Attribute VB_Name = "Update_DropDown"
Option Explicit

'Code to add new dropdown to database

Sub Add_DropDown(DropDown As String, DropDownValue As String) 'Updated for Version 2.0
    
    Application.EnableCancelKey = xlDisabled
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim qry As String
        
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
        
    qry = "Select * from tblDropDown"
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    With rst
       
        .AddNew
        .Fields("Field_Name").Value = DropDown
        .Fields("Drop_Down").Value = DropDownValue
        .Fields("Created_By").Value = ['Login Details'!$B$2]
        .Fields("Create_On").Value = [Now()]
        .Fields("Modified_By").Value = ['Login Details'!$B$2]
        .Fields("Modified_On").Value = [Now()]
        
        
        .Update
        
        .Close
        
    End With
    
    cnn.Close
    
    MsgBox "New Activity Code has been added successfully.", vbInformation, "Activity Code"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
End Sub



