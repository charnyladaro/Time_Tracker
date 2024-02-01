Attribute VB_Name = "Form_Call"
Option Explicit

'Code to Load Form - This code is being utilized on assigning macro on Home Sheet

Public Sub Load_Form()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    frmHome.Show
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub
