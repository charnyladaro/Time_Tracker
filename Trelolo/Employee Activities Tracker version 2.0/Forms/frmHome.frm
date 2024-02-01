VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHome 
   Caption         =   "Home"
   ClientHeight    =   3072
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10512
   OleObjectBlob   =   "frmHome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnLogout_Click()
    
    Dim iConfirmation As VbMsgBoxResult
    
    iConfirmation = MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion, "Logout")
    
    If iConfirmation = vbNo Then Exit Sub
    
    'Reset the Login and other support details
    
    ThisWorkbook.Sheets("Login Details").Range("A2:D" & Application.Rows.Count).Value = ""
    ThisWorkbook.Sheets("RawData").Range("A2:J" & Application.Rows.Count).Value = ""
    ThisWorkbook.Sheets("UM_Support").Range("A2:F" & Application.Rows.Count).Value = ""
            
    On Error Resume Next
    
    'Closing all forms
    
    'Unload User_Management_frm
    Unload Reporting_frm
    Unload MyCalendar
    Unload Manage_DropDown_Frm
    Unload ActivityTracker
    Unload frmHome
    Unload frmLogin
    
    MsgBox "You have successfully logged out!", vbOKOnly + vbInformation, "Logged out"

End Sub



Private Sub cmdDatabasePath_Click()
    
    Dim rng As Range
    Dim msgValue As VbMsgBoxResult
    Dim sDatabasePath As String
    
    Set rng = ThisWorkbook.Sheets("Database Path").Range("A2")
    
    
    If Trim(rng.Value) <> "" And Dir(rng.Value) <> "" Then
    
        msgValue = MsgBox("Database path has already been set. Do you want to change the Database path?", vbYesNo + vbQuestion, "Database Path")
        
        If msgValue = vbNo Then Exit Sub
   Else
   
    msgValue = MsgBox("Do you want to change the Database path?", vbYesNo + vbQuestion, "Database Path")
        
        If msgValue = vbNo Then Exit Sub
        
   End If
   
   Load frmDatabase
   
   frmDatabase.txtPath.Value = IIf(Trim(rng.Value) = "", "Default", Trim(rng.Value))
   
   frmDatabase.Show
    
    
    
End Sub

'Showing Welocome Message as lable on Home form

Private Sub UserForm_Activate() ' Code Updated for version 2.0
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
        
    Dim sUserName As String
    
    
    With ThisWorkbook.Sheets("Login Details")
       
        Me.lbl_UserID.Caption = .Range("A2").Value
        sUserName = .Range("B2").Value
        Me.lbl_role.Caption = VBA.UCase(.Range("D2").Value)
        Me.lbl_role.ControlTipText = "User's role : " & VBA.UCase(.Range("D2").Value)
        
    End With
    
    'Disabling the Drop Down button for normal user
    If Me.lbl_role.Caption <> "ADMIN" Then
    
        Me.lbl_role.Left = 320
        Me.btnLogout.Left = 368
        Me.Width = 406
        Me.btn_Manage_Dropdown.Enabled = False
        Me.btn_Manage_Dropdown.Visible = False
        Me.Label53.Visible = False
        Me.cmdDatabasePath.Visible = False
    
    End If
    
    

    If VBA.Time <= VBA.TimeValue("23:59:59") Then
        Me.lbl_welcome_msg.Caption = "Good Evening, " & sUserName
    End If
    
    If VBA.Time <= VBA.TimeValue("17:00:00") Then
        Me.lbl_welcome_msg.Caption = "Good Afternoon, " & sUserName
    End If
    
    If VBA.Time <= VBA.TimeValue("12:00:00") Then
        Me.lbl_welcome_msg.Caption = "Good Morning, " & sUserName
    End If
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
  
End Sub

Private Sub btn_ActivityTracker_Click() ' Code Updated for version 2.0
              
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Me.Hide
    Unload Me
    DoEvents
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
    ActivityTracker.Show
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub


Private Sub btn_dashboard_Click() ' Code Updated for version 2.0
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
       
    Dim a As VbMsgBoxResult
    
    a = MsgBox("Do you want to export the activitie(s) log in Excel file?", vbYesNo + vbQuestion, "Export")
    
    If a = vbNo Then Exit Sub
    
    
    'Me.Hide
    Unload Me
    DoEvents
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
    Reporting_frm1.Show
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub



Private Sub btn_Manage_Dropdown_Click() ' Code Updated for version 2.0
    
    
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Manage_DropDown_Frm.Show
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Private Sub btn_user_management_Click()


On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    User_Management_frm.Show
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    
    If CloseMode = vbFormControlMenu Then
    
        MsgBox "Please use the Logout button to close the application?", vbOKOnly + vbInformation, "Close"
    
        Cancel = True
        
    End If
    
End Sub
