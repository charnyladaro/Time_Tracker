Attribute VB_Name = "Create_Update_User_Module"
Option Explicit

Sub Create_Update_User(new_user As Boolean)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo err_msg
    
    '================ Validation ===============
    If User_Management_frm.txt_user_id.Value = "" Then
        MsgBox "Please enter the user id.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
    
    If User_Management_frm.txt_User_Name.Value = "" Then
        MsgBox "Please enter the user Name.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
    
    If User_Management_frm.cmb_role.Value = "" Then
        MsgBox "Please select the role.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
    
    If User_Management_frm.txt_supervisor.Value = "" Then
        MsgBox "Please enter the Supervisor's name.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
     
    If User_Management_frm.txt_Password.Value = "" Then
        MsgBox "Please enter a valid password.", vbCritical, "Activities Tracker"
        Exit Sub
    End If
    
    '====================================================
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim Id As String
    Dim qry As String, i As Integer
    Dim n As Long
     
    '***********************************************
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    
    If new_user = False Then
    qry = "Select * from tblUserManagment WHERE User_Id='" & User_Management_frm.txt_user_id.Value & "'"
    Else
    qry = "Select * from tblUserManagment WHERE User_Id='" & User_Management_frm.txt_user_id.Value & "' OR User_Name='" & User_Management_frm.txt_User_Name.Value & "'"
    End If
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    If new_user = False Then
        If rst.RecordCount = 0 Then
           MsgBox "No data found for this user id.", vbCritical, "Activities Tracker"
           rst.Close
           cnn.Close
           Exit Sub
        End If
    Else
    
        If rst.RecordCount > 0 Then
           MsgBox "User is already available in database", vbCritical, "Activities Tracker"
           rst.Close
           cnn.Close
           Exit Sub
        End If
    End If
    
    If new_user = True Then
        rst.AddNew
        rst.Fields("User_Id").Value = User_Management_frm.txt_user_id.Value
    End If
    
    rst.Fields("User_Name").Value = User_Management_frm.txt_User_Name.Value
    rst.Fields("Supervisor").Value = User_Management_frm.txt_supervisor.Value
    rst.Fields("Role").Value = User_Management_frm.cmb_role.Value
    
    rst.Fields("Password").Value = User_Management_frm.txt_Password.Value
    
    rst.Fields("Created_by").Value = frmHome.lbl_UserID.Caption
    rst.Fields("Modified_by").Value = frmHome.lbl_UserID.Caption
    rst.Fields("Created_on").Value = VBA.Now
    rst.Fields("Modified_on").Value = VBA.Now
    
    rst.Update
    
    rst.Close
    cnn.Close
    
    User_Management_frm.txt_Password.Value = ""
    User_Management_frm.txt_user_id.Value = ""
    User_Management_frm.txt_User_Name.Value = ""
    User_Management_frm.txt_supervisor.Value = ""
    
    
    If VBA.UCase(frmHome.lbl_role.Caption) = "ADMIN" Then
        User_Management_frm.btn_submit.Default = True
        User_Management_frm.btn_submit.Enabled = True
        User_Management_frm.btn_Update.Enabled = False
    End If
     
    With User_Management_frm.cmb_role
        .Clear
        .AddItem "ADMIN"
        .AddItem "USER"
    End With
    
    Call UM_List_box_Data 'Refreshing ListBox on User Management Form
   
    If new_user = True Then
        MsgBox "New User has been created successfully.", vbInformation, "Activities Tracker"
    Else
        MsgBox "User's details has been updated successfully.", vbInformation, "Activities Tracker"
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
      
    Exit Sub
    
err_msg:
    
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

