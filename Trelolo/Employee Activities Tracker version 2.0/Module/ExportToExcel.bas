Attribute VB_Name = "ExportToExcel"
Option Explicit

'Prepare Excel Sheet and Pasting Raw Data to that

Sub ExportDataToExcel()
 
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'On Error GoTo ErrorHandler
    
    Dim sh As Worksheet
    
    Set sh = ThisWorkbook.Sheets("RawData")
    
    If sh.Range("A" & Application.Rows.Count).End(xlUp).Row < 2 Then
        MsgBox "No Data Found", vbCritical, "Export To Excel"
        Exit Sub
    End If
    
   Dim wkb As Workbook
   Dim nsh As Worksheet
    
   Set wkb = Workbooks.Add
   Set nsh = wkb.Sheets(1)

   sh.UsedRange.Copy nsh.Range("A1")
    
    With nsh.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlJustify
        .EntireColumn.ColumnWidth = 18
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
     
    ActiveWindow.DisplayGridlines = False
     
    nsh.Name = "Data"
    Reporting_frm.Hide
    MsgBox "Data has been successfully exported in Excel.", vbInformation, "Export to Excel"
    wkb.Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    
   
ErrorHandler:

  MsgBox Err.Description, vbCritical, "Error"

  Application.DisplayAlerts = True
  Application.ScreenUpdating = True

End Sub


