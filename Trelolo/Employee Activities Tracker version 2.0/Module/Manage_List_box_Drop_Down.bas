Attribute VB_Name = "Manage_List_box_Drop_Down"
Option Explicit

Sub Manage_List_box_Data() ' Updated for version 2.0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    'On Error GoTo err_msg
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Drop_Down_Details")
    
    sh.Cells.ClearContents
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String, i As Integer
    Dim n As Long
    
        
    If Manage_DropDown_Frm.cmbDropDown.Value = "ALL" Then
    
        If Manage_DropDown_Frm.TextBox2.Value = "" Or Manage_DropDown_Frm.TextBox2.Value = "Enter Search Keyword" Then
            qry = "Select * from tblDropDown"
        Else
            qry = "Select * from tblDropDown Where Drop_Down LIKE '%" & Manage_DropDown_Frm.TextBox2.Value & "%';"
        End If
    
    Else
    
        If Manage_DropDown_Frm.TextBox2.Value = "" Or Manage_DropDown_Frm.TextBox2.Value = "Enter Search Keyword" Then
            qry = "Select * from tblDropDown Where Field_Name = '" & Manage_DropDown_Frm.cmbDropDown.Value & "'"
        Else
            qry = "Select * from tblDropDown Where Field_Name = '" & Manage_DropDown_Frm.cmbDropDown.Value & "' and Drop_Down LIKE '%" & Manage_DropDown_Frm.TextBox2.Value & "%';"
        End If
    
    End If
    
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("a2").CopyFromRecordset rst
    
   
    'n = rst.Fields.Count
      
    For i = 1 To rst.Fields.Count
        sh.Cells(1, i).Value = rst.Fields(i - 1).Name
    Next i
    
            
    rst.Close
    cnn.Close
    
    sh.Range("A:A").NumberFormat = "0"
    
    '*****************************************************
    
    Manage_DropDown_Frm.ListBox1.ColumnCount = 5
    Manage_DropDown_Frm.ListBox1.ColumnHeads = True
    Manage_DropDown_Frm.ListBox1.ColumnWidths = "0,150,150,0,0"
    
    
    ThisWorkbook.Activate
    
    n = sh.Range("A" & Application.Rows.Count).End(xlUp).Row
    
    If n > 1 Then
        Manage_DropDown_Frm.ListBox1.RowSource = "Drop_Down_Details!A2:E" & n
    Else
        Manage_DropDown_Frm.ListBox1.RowSource = "Drop_Down_Details!A2:E2"
    End If
    
    If (n - 1) < 2 Then
            Manage_DropDown_Frm.count_lbl.Caption = (n - 1) & " Item"
        ElseIf (n - 1) > 1 Then
         Manage_DropDown_Frm.count_lbl.Caption = (n - 1) & " Items"
    End If
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
err_msg:
    
        
   MsgBox "Error Number:-" & Err.Number & vbLf & vbLf & " || Error Detail:- " & Err.Description & " || "
    
   Application.DisplayAlerts = True
   Application.ScreenUpdating = True


End Sub




