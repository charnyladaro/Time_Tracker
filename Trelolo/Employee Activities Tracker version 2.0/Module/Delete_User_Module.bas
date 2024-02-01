Attribute VB_Name = "Delete_User_Module"
Option Explicit


Sub Delete_User()

    'Application.EnableCancelKey = xlDisabled
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo err_msg
    
    '================ confirmation ===============
    Dim iInput As Integer
    
    iInput = MsgBox("Do you want to delete selected user(s)?", vbQuestion + vbYesNo, "Activities Tracker")
    If iInput = vbNo Then Exit Sub
    '====================================================
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim qry As String
    '***********************************************
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    Dim X, n As Long
    X = 0
    With User_Management_frm.lstUserDetails
        For n = 0 To .ListCount - 1
            
            If .Selected(n) Then
                qry = "Delete * FROM tblUserManagment WHERE ID = " & .List(n, 0)
                rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
                X = X + 1
            End If
            
        Next n
    End With
    
    cnn.Close
    
    MsgBox "User has been deleted successfully.", vbInformation, "Activities Tracker"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
     
     Exit Sub
    
err_msg:
    
        MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

End Sub


