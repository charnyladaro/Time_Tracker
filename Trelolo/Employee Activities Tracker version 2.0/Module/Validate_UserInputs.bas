Attribute VB_Name = "Validate_UserInputs"
Option Explicit

'Code to reset the controls available on Activity Tracker form

Sub Reset_Control() 'Updated for Version 2.0
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler

    With ActivityTracker
    
        'Enabling
        
        .cmdLogin.Enabled = False
        .cmdLogout.Enabled = True
        .lstActivityCode.Enabled = True
        .txtDescription.Enabled = True
        .cmdStart.Enabled = True
        .cmdEnd.Enabled = False
        
        'Reseting
        .txtDescription.Value = ""
        .lstActivityCode.Value = "Select Activity Code"
    
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

'Code to Validate teh entry or selection done by user

Public Function Validation() As Boolean 'Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler

    Validation = True

    With ActivityTracker
    
    'Client Name
     
     If Trim(.cmbClientName.Value) = "Select Client Name" Or Trim(.cmbClientName.Value) = "" Then
        MsgBox "Please select the Client Name from drop-down.", vbOKOnly + vbInformation, "Incorrect Entry"
        .cmbClientName.SetFocus
        Validation = False
        Exit Function
     End If
     
     'Location
     
     If Trim(.cmbLocationName.Value) = "Select Location" Or Trim(.cmbLocationName.Value) = "" Then
        MsgBox "Please select the Location from drop-down.", vbOKOnly + vbInformation, "Incorrect Entry"
        .cmbLocationName.SetFocus
        Validation = False
        Exit Function
     End If

    'Activity Code
     
     If Trim(.lstActivityCode.Value) = "Select Activity Code" Or Trim(.lstActivityCode.Value) = "" Then
        MsgBox "Please select the Activity Code from drop-down.", vbOKOnly + vbInformation, "Incorrect Entry"
        .lstActivityCode.SetFocus
        Validation = False
        Exit Function
     End If
          
     'Description
     If Trim(.txtDescription.Value) = "" Then
        MsgBox "Please enter the description.", vbOKOnly + vbInformation, "Incorrect Entry"
        .txtDescription.SetFocus
        Validation = False
        Exit Function
     End If
     
    End With
     
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True


End Function

'Code to Lock user controls after starting any activity

Sub Lock_UserInput() 'Updated for version 2.0
   
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
   'On Error GoTo ErrorHandler

    With ActivityTracker
    
        .cmdLogout.Enabled = False
        .cmdLogin.Enabled = False
        .cmbClientName.Enabled = False
        .cmbLocationName.Enabled = False
        .lstActivityCode.Enabled = False
        .txtDescription.Enabled = False
        .cmdStart.Enabled = False
        .cmdEnd.Enabled = True
        .cmdRefresh.Enabled = False
    
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub

'Code to validate whether new drop down item is available in database or not
Function Value_Exist(DropDown As String, DropDownValue As String) As Boolean 'Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim sh As Worksheet
    Dim fRange As Range
    
    Set sh = ThisWorkbook.Sheets("Drop-Down")
    
    Select Case DropDown
        
        Case "Activity Code"
            
            Set fRange = sh.Range("A:A").Find(DropDownValue)
            If Not fRange Is Nothing Then
                Value_Exist = True
                MsgBox "Activity Code already exists in database."
                Manage_DropDown_Frm.txtValue.SetFocus
                Exit Function
            End If
        
    End Select
   
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
  
End Function


