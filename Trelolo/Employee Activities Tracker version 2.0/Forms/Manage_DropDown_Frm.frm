VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Manage_DropDown_Frm 
   Caption         =   "Manage Drop-Down"
   ClientHeight    =   7200
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7548
   OleObjectBlob   =   "Manage_DropDown_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Manage_DropDown_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbDropDown_Change() ' Done for version 2.0
      
    Me.txtValue.Value = ""
    
    If Me.cmbDropDown.Value = "ALL" Then
    
        txtValue.Enabled = False
        txtValue.ControlTipText = "Select the Drop-down type to add."
        
    Else
    
        txtValue.Enabled = True
        txtValue.ControlTipText = "Enter " & Me.cmbDropDown.Value
        lblDropDown.Caption = Me.cmbDropDown.Value
    
    End If
     
    Call Manage_List_box_Data
     
    
    
End Sub

Private Sub cmdDelete_Click() ' Done for Version 2.0
    
    Dim n, i As Long
    
    n = 0
    For i = 0 To Me.ListBox1.ListCount - 1
    
        If Me.ListBox1.Selected(i) = True Then n = n + 1
    
    Next i
    
    If n = 0 Then
        MsgBox "No Record was selected.", vbCritical, "Drop-down"
        Exit Sub
    End If
    
    If Me.ListBox1.List(Me.ListBox1.ListIndex, 0) = "" Then Exit Sub
    
    Call Delete_List

End Sub

Private Sub cmdExport_Click() ' Done for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo err_msg
    
    Dim nwb As Workbook
    Set nwb = Workbooks.Add
    Dim nsh As Worksheet
    Set nsh = nwb.Sheets(1)
    
    ThisWorkbook.Sheets("Drop_Down_Details").UsedRange.Copy nsh.Range("A1")
    
    With nsh.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlJustify
        .EntireColumn.ColumnWidth = 12
        .EntireRow.RowHeight = 15
        .Font.Size = 10
        .Font.Name = "Calibri"
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
    End With
    
    With nsh.Range("A1", nsh.Cells(1, Application.CountA(nsh.Range("1:1"))))
        .Font.Bold = True
        .Interior.ColorIndex = 15
    End With
    
    
    nsh.Activate
    nsh.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    
    
    MsgBox "Drop-down data have been exported to Excel.", vbInformation, "Export"
     
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
     
    Exit Sub
    
err_msg:
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

 
End Sub

Private Sub cmdRefresh_Click() ' Done for version 2.0
    If Me.cmbDropDown.Value <> "" Then Call Manage_List_box_Data
End Sub

Private Sub cmdReset_Click() ' Done for Version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim v As VbMsgBoxResult
    
    v = MsgBox("Do you want to reset this form?", vbYesNo + vbQuestion, "Reset")
    
    If v = vbNo Then Exit Sub
    
    Me.cmbDropDown.Value = "ALL"
    Me.txtValue.Value = ""
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
   
End Sub




Private Sub UserForm_Initialize() ' Done for Version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
    'On Error GoTo ErrorHandler

    With Me.cmbDropDown
        .Clear
        .AddItem "ALL"
        .AddItem "Activity Code"
        .AddItem "Client Name"
        .AddItem "Location"
        '.AddItem "LOB_Task"
        .Value = "ALL"
        
    End With
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

'Validation Code
Function ValidateInput() As Boolean ' Done for version 2.0
   
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ValidateInput = True
    
    If Me.cmbDropDown.Value = "ALL" Then
        
        MsgBox "Please Select Drop-down Type.", vbOKOnly + vbInformation, "Invalid Entry"
        Me.cmbDropDown.SetFocus
        ValidateInput = False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Function
       
    End If
    
    If Trim(Me.txtValue.Value) = "" Then
        
        MsgBox "Please enter new Activity code.", vbOKOnly + vbInformation, "Invalid Entry"
        Me.txtValue.SetFocus
        ValidateInput = False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Function
       
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
        
End Function



Private Sub cmdSubmit_Click()
        
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    
    'On Error GoTo ErrorHandler

    
    If Not ValidateInput Then Exit Sub
    
    
    If Value_Exist(Me.cmbDropDown, Me.txtValue) Then Exit Sub
    
    'Code to update data to the database
    '<--------------------------------->
        
    Call Add_DropDown(Trim(Me.cmbDropDown.Value), Trim(Me.txtValue.Value))
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    sh.Cells.Delete 'Clearing Previous Items
        
    Call Get_LocationList 'Prepare Location List for Activities Tracker Drop-down
    Call Get_ActivityList 'Prepare Activity Code List for Activities Tracker Drop-down
    Call Get_ClientNameList 'Prepare Client Name List for Activities Tracker Drop-down
        
    
    Me.txtValue.Value = ""
    
    Call Manage_List_box_Data
     
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub



   








