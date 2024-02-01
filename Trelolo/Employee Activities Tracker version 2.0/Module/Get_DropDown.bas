Attribute VB_Name = "Get_DropDown"
Option Explicit

'Procedure to get Drop-Down - Activity Name

Public Sub Get_ActivityList()
    
    
    Application.EnableCancelKey = xlDisabled
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    'sh.Cells.Delete
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String
    
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    sh.Range("A1").Value = "Activity Code"
    
    qry = "SELECT [Drop_Down] from tblDropDown where [Field_Name]='Activity Code'"
        
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("A3").CopyFromRecordset rst
            
    rst.Close
    
    cnn.Close
    
    'Sorting Fields
    sh.Sort.SortFields.Clear
    sh.Range("A:A").Sort Key1:=sh.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    
    sh.Range("A2").Insert Shift:=xlDown
    sh.Range("A2").Value = "Select Activity Code"
    
      
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub


'Procedure to get Drop-Down - Location Name

Public Sub Get_LocationList()
    
    
    Application.EnableCancelKey = xlDisabled
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    'sh.Cells.Delete
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String
    
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    sh.Range("D1").Value = "Location"
    
    qry = "SELECT [Drop_Down] from tblDropDown where [Field_Name]='Location'"
        
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("D3").CopyFromRecordset rst
            
    rst.Close
    
    cnn.Close
    
    'Sorting Fields
    sh.Sort.SortFields.Clear
    sh.Range("D:D").Sort Key1:=sh.Range("D1"), Order1:=xlAscending, Header:=xlYes
    
    
    sh.Range("D2").Insert Shift:=xlDown
    sh.Range("D2").Value = "Select Location"
    
      
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub


'Procedure to get Drop-Down - Client Name

Public Sub Get_ClientNameList()
    
    
    Application.EnableCancelKey = xlDisabled
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("ListBox_Value")
    
    'sh.Cells.Delete
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String
    
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
    
    sh.Range("G1").Value = "Client Name"
    
    qry = "SELECT [Drop_Down] from tblDropDown where [Field_Name]='Client Name'"
        
    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("G3").CopyFromRecordset rst
            
    rst.Close
    
    cnn.Close
    
    'Sorting Fields
    sh.Sort.SortFields.Clear
    sh.Range("G:G").Sort Key1:=sh.Range("G1"), Order1:=xlAscending, Header:=xlYes
    
    
    sh.Range("G2").Insert Shift:=xlDown
    sh.Range("G2").Value = "Select Client Name"
    
      
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub

