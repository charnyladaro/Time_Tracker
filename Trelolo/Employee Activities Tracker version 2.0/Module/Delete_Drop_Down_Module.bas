Attribute VB_Name = "Delete_Drop_Down_Module"
Option Explicit

Sub Delete_List()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo err_msg
    
    
    Dim iConfirmation As VbMsgBoxResult
    iConfirmation = VBA.MsgBox("Do you want to permanent delete the selected records?", vbQuestion + vbYesNo, "Drop-down")
    
    
    If iConfirmation = vbNo Then Exit Sub
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim rst2 As New ADODB.Recordset
    
    Dim qry As String, n As Integer
    
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    With Manage_DropDown_Frm.ListBox1
    
        For n = 0 To .ListCount - 1
        
            If .Selected(n) Then
            
                qry = "Delete * FROM tblDropDown WHERE Serial_No = " & .List(n, 0) & ";"
                
                rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
                
                
            End If
        
        Next n
    End With
    
    
    cnn.Close
    
    Call Manage_List_box_Data ' Refreshing List box data
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
err_msg:
            
        MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True


End Sub








