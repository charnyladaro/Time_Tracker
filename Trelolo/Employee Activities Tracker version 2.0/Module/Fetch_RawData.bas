Attribute VB_Name = "Fetch_RawData"
Option Explicit


'Code to fetch data from Database and update to Raw Data Sheet

Sub FetchRawData(dStartDate As Variant, dEndDate As Variant)
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
       
    Dim i As Integer
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("RawData")
    
    sh.Cells.Delete
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim qry As String
     
    
    #If Win64 Then
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #Else
        cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabasePath & ";Jet OLEDB:Database Password=thedatalabs"
    #End If
     
    qry = "Select * From tblActivityLog WHERE Dates>=#" & dStartDate & "# AND Dates<=#" & dEndDate & "#"
    
    'Please use the below code to restrict user specifc data extraction only
    
    If VBA.UCase(['Login Details'!$D$2]) <> "ADMIN" Then      ' Replace with Admin ID

        qry = qry & " AND [Employee ID]='" & VBA.UCase(['Login Details'!$A$2]) & "'"

    End If
    
    qry = qry & " ORDER BY [Submitted On] DESC"

    rst.Open qry, cnn, adOpenKeyset, adLockOptimistic
    
    sh.Range("a2").CopyFromRecordset rst
          
    For i = 1 To rst.Fields.Count
        sh.Cells(1, i).Value = VBA.UCase(rst.Fields(i - 1).Name)
    Next i
        
    rst.Close
    cnn.Close
      
    'Applying formatting to columns
    sh.Range("B:B").NumberFormat = "D-MMM-YY"
    sh.Range("J:K").NumberFormat = "D-MMM-YY HH:MM:SS AM/PM"
    sh.Range("L:L").NumberFormat = "HH:MM:SS"
    sh.Range("M:M").NumberFormat = "D-MMM-YY HH:MM:SS AM/PM"
    
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub
    

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True


End Sub

