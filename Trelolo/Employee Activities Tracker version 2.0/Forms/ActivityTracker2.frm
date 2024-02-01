VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ActivityTracker 
   Caption         =   "Activity Logger"
   ClientHeight    =   10320
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11484
   OleObjectBlob   =   "ActivityTracker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ActivityTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtr _
        Lib "user32.dll" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long) As LongPtr

    Private Declare PtrSafe Function SetWindowLongPtr _
        Lib "user32.dll" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) As LongPtr

    Private Declare PtrSafe Function FindWindowA _
        Lib "user32.dll" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar _
        Lib "user32.dll" ( _
        ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function GetWindowLongPtr _
        Lib "user32.dll" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

    Private Declare Function SetWindowLongPtr _
        Lib "user32.dll" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

    Private Declare Function FindWindowA _
        Lib "user32.dll" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

    Private Declare Function DrawMenuBar _
        Lib "user32.dll" ( _
        ByVal hwnd As Long) As Long
#End If


'Code to Add Minimize and Maximizie Button To Form
Private Sub CreateMenu()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'On Error GoTo ErrorHandler

    Const GWL_STYLE As Long = -16
    Const WS_SYSMENU As Long = &H80000
    Const WS_MINIMIZEBOX As Long = &H20000
    'Const WS_MAXIMIZEBOX As Long = &H10000 ' Maximize button disabled

    #If Win64 Then
        Dim lngFrmWndHdl As LongPtr
        Dim lngStyle As LongPtr
    #Else
        Dim lngFrmWndHdl As Long
        Dim lngStyle As Long
    #End If

    lngFrmWndHdl = FindWindowA(vbNullString, Me.Caption)

    lngStyle = GetWindowLongPtr(lngFrmWndHdl, GWL_STYLE)
    lngStyle = lngStyle Or WS_SYSMENU       'Add SystemMenu
    lngStyle = lngStyle Or WS_MINIMIZEBOX   'Add MinimizeBox
    'lngStyle = lngStyle Or WS_MAXIMIZEBOX   'Add MaximizeBox

    SetWindowLongPtr lngFrmWndHdl, GWL_STYLE, lngStyle

    DrawMenuBar lngFrmWndHdl

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error"

  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub
Private Sub cmdClose_Click()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler

    If Me.cmdEnd.Enabled = True And Me.cmdLogin.Enabled = False And Me.cmdLogout.Enabled = False Then
    
        MsgBox "Please click on 'End' button to stop the current activity before closing this.", vbOKOnly + vbCritical, "Close"
        
        Exit Sub
        
    End If
    
    If Me.cmdLogin.Enabled = False And Me.cmdLogout.Enabled = True Then
    
        MsgBox "Please click on 'Day Out' button before closing this.", vbOKOnly + vbCritical, "Close"
        
        Exit Sub
        
    End If
   
    Unload ActivityTracker
    DoEvents
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
    frmHome.Show
    DoEvents
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
        
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    
End Sub






Private Sub lblTheDataLabs_Click()

    Dim URL As String
    URL = "https://www.thedatalabs.org"
    ActiveWorkbook.FollowHyperlink URL
    
End Sub



Private Sub UserForm_Initialize()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Call CreateMenu                 ' Creating Minimize, Maximize on Form
    Call Disable_Activity_Control   ' Disabling controls
    Call Get_ActivityLog            ' Fetching activity data from Database
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    sh.Cells.Delete 'Clearing Previous Items
    
    Call Get_ClientNameList         ' Fetching Client Name List
    Call Get_LocationList           ' Fetching Location List
    Call Get_ActivityList           ' Fetching Activity List
        
    Call DropDownRange              ' Assigned Drop-Down Range to Location, Client and Activity
    
    Me.ActivityHeader = "Activities Log: "
    Me.lblDate = VBA.Format(Date, "DD MMMM YYYY")
    Me.lblCount = "# records : " & Me.lstActivity.ListCount
    
    cmdLogin.Enabled = True
    cmdLogout.Enabled = False
    Me.txtDate = VBA.Format(Date, "DD-MMMM-YYYY")
    
    'Assiging the User Details as per Login
    
    Me.txtEmployeeID.Value = ['Login Details'!$A$2]
    Me.txtEmployeeName.Value = ['Login Details'!$B$2]
    Me.txtSupervisor.Value = ['Login Details'!$C$2]
    
    'If unexpected logged out, load the last activity
    Call Load_Incompleted_Log
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

'Creating Maximize and Minimize Menu on Form Control
Private Sub UserForm_Activate() ' Updated for version 2.0
    Call CreateMenu
End Sub

'Code for Day In

Private Sub cmdLogin_Click() ' Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Me.txtDate = TodayDate()                             ' Date user defined function
    
    Call Enable_Activity_Control                         ' Enabling all the controls
    
    Me.cmdLogin.Enabled = False
    Me.cmdLogout.Enabled = True
    
    FirstLoginTime = Now()
    Me.LoginTime.Caption = Format(FirstLoginTime, "HH:MM:SS AM/PM")
    
    TimerActive = True                                  'Start time
    
    Call Timer
    
    Call Add_StartEntry("Login", True, FirstLoginTime)
    
    'Assiging the User Details as per Login
    
    Me.txtEmployeeID.Value = ['Login Details'!$A$2]
    Me.txtEmployeeName.Value = ['Login Details'!$B$2]
    Me.txtSupervisor.Value = ['Login Details'!$C$2]
    
    Me.lblCount = "# records : " & Me.lstActivity.ListCount - 1
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
    
ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
     
End Sub

'Code for Day Out

Private Sub cmdLogout_Click() ' Updated for version 2.0
            
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
        
    Me.txtDate = TodayDate()                   ' Date user defined function
        
    Dim a As VbMsgBoxResult
    
    a = MsgBox("Are you sure to stop hours tracking?", vbYesNo + vbQuestion, "Logout")
    
    If a = vbNo Then Exit Sub
        
    TimerActive = False 'End Time
    
    Call Add_StartEntry("Login", False, Now()) ' Checked
    
    Call Disable_Activity_Control              ' Disabling controls after logout
    
    Me.cmdLogin.Enabled = True
    Me.cmdLogout.Enabled = False
    
    FirstLoginTime = 0
    
    'Assiging the User Details as per Login
    
    Me.txtEmployeeID.Value = ['Login Details'!$A$2]
    Me.txtEmployeeName.Value = ['Login Details'!$B$2]
    Me.txtSupervisor.Value = ['Login Details'!$C$2]
    
    Me.lblCount = "# records : " & Me.lstActivity.ListCount
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
    
End Sub


'Code to Start Activity after login

Private Sub cmdStart_Click() ' Updated for Version 2.0
           
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Me.txtDate = TodayDate()                                                'Date user defined function
    
    On Error GoTo ErrorHandler
 
 If Me.lstActivityCode.Value = "Break" Or Me.lstActivityCode.Value = "15 Minutes Break" _
  Or Me.lstActivityCode.Value = "30 Minutes Break" Or Me.lstActivityCode.Value = "Lunch Break" Then                                  'If Task is Break then no validation requried
    
    CurrentActivityTime = Now()
    Call Add_StartEntry(Me.lstActivityCode.Value, True, CurrentActivityTime) 'Checked
    Call Lock_UserInput
 
 ElseIf Validation = True And (Me.lstActivityCode.Value <> "Break" Or Me.lstActivityCode.Value <> "15 Minutes Break" _
  Or Me.lstActivityCode.Value <> "30 Minutes Break" Or Me.lstActivityCode.Value <> "Lunch Break") Then     'If task is not break then validation required.
    
    CurrentActivityTime = Now()
    Call Add_StartEntry(Me.lstActivityCode.Value, True, CurrentActivityTime)
    Call Lock_UserInput
 
 Else
 
    Exit Sub
    
 End If
    
    Me.txtEmployeeID.Value = ['Login Details'!$A$2]
    Me.txtEmployeeName.Value = ['Login Details'!$B$2]
    Me.txtSupervisor.Value = ['Login Details'!$C$2]
    
    Me.lblCount = "# records : " & Me.lstActivity.ListCount
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
 
End Sub

'Code to Stop ongoing activity

Private Sub cmdEnd_Click() ' Updated for Version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Me.txtDate = TodayDate()                                   'Date user defined function
    
    Dim a As VbMsgBoxResult
    
    a = MsgBox("Are you sure you want to stop the current activity?", vbYesNo + vbQuestion, "Stop Activity")
    
    If a = vbNo Then Exit Sub
    
    CurrentActivityTime = 0
    
    Call Add_StartEntry(Me.lstActivityCode.Value, False, Now()) ' Checked
    
    'Enabling Fileds
    
     With Me
        .cmdLogin.Enabled = False
        .cmdLogout.Enabled = True
        .lstActivityCode.Enabled = True
        .cmbClientName.Enabled = True
        .cmbLocationName.Enabled = True
        .txtDescription.Enabled = True
        .cmdStart.Enabled = True
        .cmdEnd.Enabled = False
        .txtDescription.Value = ""
        .cmdRefresh.Enabled = True
    End With
    
    
    ActivityTracker.CurrentActivityHours.Caption = "00:00:00"
    
    Me.txtEmployeeID = VBA.UCase(Environ("Username"))
    
    Me.lblCount = "# records : " & Me.lstActivity.ListCount
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
     
End Sub

'Code to Export raw data to Excel file

Private Sub Export_Click() ' Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
   
    'On Error GoTo ErrorHandler
    
    Dim a As VbMsgBoxResult
    
    a = MsgBox("Do you want to export the activitie(s) log in Excel file?", vbYesNo + vbQuestion, "Export to Excel")
    
    If a = vbNo Then Exit Sub
    
    Reporting_frm.Show                  ' Open form to select Date range
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
    
End Sub

'Code to Disable Controls after logout
Sub Disable_Activity_Control()  ' Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    With Me
    
        .cmdLogin.Enabled = True
        .cmdLogout.Enabled = False
        .cmbClientName.Enabled = False
        .cmbLocationName.Enabled = False
        .lstActivityCode.Enabled = False
        .txtDescription.Enabled = False
        .cmdStart.Enabled = False
        .cmdEnd.Enabled = False
    
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

'Code to Enable Controls after login
Sub Enable_Activity_Control()   ' Updated for Version 2.0
       
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler

    With Me
        
        .cmdLogin.Enabled = False
        .cmdLogout.Enabled = True
        .cmbClientName.Enabled = True
        .cmbLocationName.Enabled = True
        .lstActivityCode.Enabled = True
        .txtDescription.Enabled = True
        .cmdStart.Enabled = True
        .cmdEnd.Enabled = False
    
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

'Code to assign Acivities name to ComboBox

Sub DropDownRange()  ' Updated for Drop down range
   
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
     'Activity Name
    
    On Error Resume Next
    
    
    
    'Activity List
    Me.lstActivityCode.Clear
    
    'On Error GoTo ErrorHandler
    
    If sh.Range("A" & Application.Rows.Count).End(xlUp).Row > 1 Then
    
        Me.lstActivityCode.RowSource = "'ListBox_Value'!A2:A" & sh.Range("A" & Application.Rows.Count).End(xlUp).Row
        Me.lstActivityCode.Value = "Select Activity Code"
    
    Else
    
       Me.lstActivityCode.RowSource = "'ListBox_Value'!A2"
       Me.lstActivityCode.Value = "Select Activity Code"
    
    End If
    
    '------------------------------------------------------------
    
    'Location Name
    Me.cmbLocationName.Clear
    
    'On Error GoTo ErrorHandler
    
    If sh.Range("D" & Application.Rows.Count).End(xlUp).Row > 1 Then
    
        Me.cmbLocationName.RowSource = "'ListBox_Value'!D2:D" & sh.Range("D" & Application.Rows.Count).End(xlUp).Row
        Me.cmbLocationName.Value = "Select Location"
    
    Else
    
       Me.lstActivityCode.RowSource = "'ListBox_Value'!D2"
       Me.lstActivityCode.Value = "Select Location"
    
    End If
    
    '-------------------------------------------------------------
    
    'Client Name
    Me.cmbClientName.Clear
    
    'On Error GoTo ErrorHandler
    
    If sh.Range("G" & Application.Rows.Count).End(xlUp).Row > 1 Then
    
        Me.cmbClientName.RowSource = "'ListBox_Value'!G2:G" & sh.Range("G" & Application.Rows.Count).End(xlUp).Row
        Me.cmbClientName.Value = "Select Client Name"
    
    Else
    
       Me.lstActivityCode.RowSource = "'ListBox_Value'!G2"
       Me.lstActivityCode.Value = "Select Client Name"
    
    End If
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
        
    
End Sub

'Code to update Combobox if any addition or deletion done in Database

Private Sub cmdRefresh_Click() ' Updated for Version 2.0
       
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    'Getting Database Path
    
        Dim sPath As String
   
         sPath = Trim(ThisWorkbook.Sheets("Database Path").Range("A2").Value)
          
         If sPath = "Default" Then
         
          sPath = ThisWorkbook.Path & "\Database\Database.mdb"
          
         End If
            
         If Dir(sPath) = "" Then
          
              MsgBox "Database file is missing. Unable to proceed.", vbOKOnly + vbCritical, "Error"
              Exit Sub
              ThisWorkbook.Close
        End If
        
        sDatabasePath = sPath
          
    
    '###############################################################
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    sh.Cells.Delete 'Clearing Previous Items
    
    Call Get_ClientNameList         ' Fetching Client Name List
    Call Get_LocationList           ' Fetching Location List
    Call Get_ActivityList           ' Fetching Activity List
    
    Call DropDownRange                  'Update the Combox Box rowsource
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
  
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


'Code to restrict closing the form if there is an active login

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer) ' Udpated for Version 2.0
    
    
    If CloseMode = vbFormControlMenu Then
    
        MsgBox "Please use the Close button available on this form?", vbOKOnly + vbInformation, "Close"
    
        Cancel = True
        
    End If
    
End Sub




