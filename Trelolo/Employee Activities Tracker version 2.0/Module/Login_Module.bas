Attribute VB_Name = "Login_Module"
Option Explicit
Global mSubject As String
Global cSubject As String

Sub Login() ' Updated for version 2.0

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo err_msg
    
    'Assigning Default Back color to User Name and Password Field
    frmLogin.txtUserID.BackColor = vbWhite
    frmLogin.txtPassword.BackColor = vbWhite
    
    If frmLogin.txtUserID.Value = "" Then
        MsgBox "Please enter the User ID.", vbInformation, "Login"
        frmLogin.txtUserID.BackColor = vbRed
        Exit Sub
    End If
    
    If frmLogin.txtPassword.Value = "" Then
        MsgBox "Please enter the Password.", vbInformation, "Login"
        frmLogin.txtPassword.BackColor = vbRed
        Exit Sub
    End If
        
    Dim str As String
    str = frmLogin.txtPassword.Value
        
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
        
    Dim qry As String
        
    qry = "SELECT * FROM tblUserManagment WHERE User_Id = '" & frmLogin.txtUserID.Value & "'"
        
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
        
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
        
    If rst.RecordCount = 0 Then
    
       MsgBox "Incorrect User Id.", vbCritical, "Login"
       frmLogin.txtUserID.Value = ""
       frmLogin.txtUserID.BackColor = vbRed
       frmLogin.txtUserID.SetFocus
       
    ElseIf rst.Fields("Password") = str Then
        
       With ThisWorkbook.Sheets("Login Details")
       
        .Range("A2").Value = rst.Fields("User_Id").Value
        .Range("B2").Value = rst.Fields("User_Name").Value
        .Range("C2").Value = rst.Fields("Supervisor").Value
        .Range("D2").Value = rst.Fields("Role").Value
        
       End With
       
       frmLogin.Hide
       frmHome.Show
       
               
   Else
   
       MsgBox "Incorrect password.", vbCritical, "Login"
       frmLogin.txtPassword.BackColor = vbRed
       frmLogin.txtPassword.Value = ""
       frmLogin.txtPassword.SetFocus
       
   End If
        
         
    cnn.Close
    
   
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    
    Exit Sub
    
err_msg:
    
   
    MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


End Sub

