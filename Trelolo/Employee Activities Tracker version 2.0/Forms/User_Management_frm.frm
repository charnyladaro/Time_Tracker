VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User_Management_frm 
   Caption         =   "User's Profile Management"
   ClientHeight    =   7416
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11208
   OleObjectBlob   =   "User_Management_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "User_Management_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Clear_Click() ' UPDATED for Activities Tracker

    On Error GoTo err_msg
    
    Me.txt_Password.Value = ""
    Me.txt_user_id.Value = ""
    Me.txt_User_Name.Value = ""
    Me.txt_supervisor.Value = ""
     
    With Me.cmb_role
        .Clear
        .AddItem "ADMIN"
        .AddItem "USER"
    End With
    
    Me.btn_submit.Default = True
    Me.btn_submit.Enabled = True
    Me.btn_Update.Enabled = False
     
    Exit Sub
    
err_msg:

    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "


End Sub


Private Sub cmd_Delete_Click() ' UPDATED for Activities Tracker

    On Error GoTo err_msg
    
    Dim n, i As Long
    n = 0
    For i = 0 To Me.lstUserDetails.ListCount - 1
    If Me.lstUserDetails.Selected(i) = True Then n = n + 1
    Next i
    
    If n = 0 Or Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 0) = "" Then
        MsgBox "Please a record to delete.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
     
    Call Delete_User
    
    Call UM_List_box_Data
    
    Exit Sub
    
err_msg:
    
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    

End Sub

Private Sub btn_submit_Click()

    On Error GoTo err_msg
    
    Call Create_Update_User(True)
    
    Exit Sub
    
err_msg:
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "

End Sub
 
Private Sub btn_Update_Click()

    On Error GoTo err_msg
    
    Call Create_Update_User(False)
         
    Exit Sub
    
err_msg:

    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
    
End Sub



Private Sub cmd_Extract_Click()

    'Application.EnableCancelKey = xlDisabled
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo err_msg
    
    Dim nwb As Workbook
    Set nwb = Workbooks.Add
    Dim nsh As Worksheet
    Set nsh = nwb.Sheets(1)
    
    ThisWorkbook.Sheets("UM_Support").Range("G:G").Delete
    
    ThisWorkbook.Sheets("UM_Support").UsedRange.Copy nsh.Range("A1")
    
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
    
    nsh.Range("F:CZ").Delete
    nsh.Range("A:A").Delete
    
    nsh.Activate
    nsh.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
err_msg:
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub



Private Sub lstUserDetails_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo err_msg
    
    On Error Resume Next
    If Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 0) <> "" Then
    
        Me.txt_user_id.Value = Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 1)
        Me.txt_User_Name.Value = Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 2)
        Me.txt_supervisor.Value = Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 3)
        Me.cmb_role.Value = Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 4)
        Me.txt_Password.Value = Me.lstUserDetails.List(Me.lstUserDetails.ListIndex, 5)
        Me.btn_submit.Enabled = False
        Me.btn_Update.Enabled = True
        Me.btn_Update.Default = True
        
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
err_msg:
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    

End Sub

Private Sub UserForm_Activate()
 
    On Error GoTo err_msg
     
    With Me.cmb_role
        .Clear
        .AddItem "ADMIN"
        .AddItem "USER"
    End With
    
    
    If VBA.UCase(frmHome.lbl_role.Caption) <> "ADMIN" Then
        Me.btn_clear.Visible = False
        Me.btn_submit.Visible = False
        Me.btn_Update.Enabled = True
        Me.btn_Update.Default = True
        Me.Height = 165
    Else
        Me.btn_submit.Default = True
        Me.btn_submit.Enabled = True
        Me.btn_Update.Enabled = False
    
    End If
     
    Call UM_List_box_Data
        
    Exit Sub
    
err_msg:

    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
End Sub

 

