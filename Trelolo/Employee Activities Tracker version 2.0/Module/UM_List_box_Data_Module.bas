Attribute VB_Name = "UM_List_box_Data_Module"
Option Explicit


'Sub routine to load data to the User Management Form along with List box

Sub UM_List_box_Data() 'Updated for version 2.0

    'Application.EnableCancelKey = xlDisabled
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo err_msg
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("UM_Support")
    
    sh.Cells.ClearContents
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String, i As Integer
    Dim n As Long
    
    
        If VBA.UCase(frmHome.lbl_role.Caption) = "ADMIN" Then
            qry = "Select ID, User_Id, User_Name,Supervisor,Role,Password from tblUserManagment"
        Else
            qry = "Select ID, User_Id, User_Name,Supervisor,Role,Password from tblUserManagment where USER_ID = '" & frmHome.lbl_UserID.Caption & "'"
        End If
    
    
    '***********************************************
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("A2").CopyFromRecordset rst
     
      
    For i = 1 To rst.Fields.Count
        sh.Cells(1, i).Value = rst.Fields(i - 1).Name
    Next i
        
    rst.Close
    cnn.Close
    
     
    sh.Range("A:A").NumberFormat = "0"
    
    
    '*****************************************************
    With User_Management_frm.lstUserDetails
        .ColumnCount = 6
        .ColumnHeads = True
        .ColumnWidths = "0,120,120,120,120,0"
     
     
     ThisWorkbook.Activate
     
    
    n = sh.Range("A" & Application.Rows.Count).End(xlUp).Row
    
    If n > 1 Then
     .RowSource = "UM_Support!A2:F" & n
    Else
     .RowSource = "UM_Support!A2:F2"
    End If
    
    End With
    
    If (n - 1) < 2 Then
        User_Management_frm.lbl_record_count.Caption = (n - 1) & " Item"
    ElseIf (n - 1) > 1 Then
        User_Management_frm.lbl_record_count.Caption = (n - 1) & " Items"
    End If
    
    
    If VBA.UCase(frmHome.lbl_role.Caption) <> "ADMIN" Then
        If n = 2 Then
            With User_Management_frm
            
                'Assigning the User Login Details
                .txt_user_id.Value = sh.Range("B2").Value
                .txt_User_Name.Value = sh.Range("C2").Value
                .txt_supervisor.Value = sh.Range("D2").Value
                .cmb_role.Value = sh.Range("E2").Value
                .txt_Password.Value = sh.Range("F2").Value
                
                'Disabling the Controls so that User can only change password
                .txt_user_id.Enabled = False
                .txt_User_Name.Enabled = False
                .txt_supervisor.Enabled = False
                .cmb_role.Enabled = False
                            
            End With
        End If
     End If
    
    
    '****************************************************
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    
err_msg:
    
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub



