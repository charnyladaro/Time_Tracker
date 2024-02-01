Attribute VB_Name = "Greeting_Message"
'Code to get the full User name form Windows operating System and displaying on Home Form

'Please use this function to get the User Full Name if you don't want to use the User Id in Welcome Messate

Function Userfullname_Windows() As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim WSHnet As Object
    
    Set WSHnet = CreateObject("WScript.Network")
    
    Set objUser = GetObject("WinNT://" & WSHnet.UserDomain & "/" & WSHnet.UserName & ",user")
    Userfullname = objUser.FullName
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
End Function


