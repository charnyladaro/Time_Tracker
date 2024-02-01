VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Reporting_frm1 
   Caption         =   "Export to Excel"
   ClientHeight    =   2616
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4908
   OleObjectBlob   =   "Reporting_frm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Reporting_frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Cal1_Click()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ''On Error GoTo ErrorHandler
    
   Call MyCalendar.DatePicker(Me.TextBox1)
   
   If CDate(Me.TextBox1.Value) > Date Then
   
        MsgBox "Start Date can't be future date.", vbOKOnly + vbInformation, "Start Date"
        Call MyCalendar.DatePicker(Me.TextBox1)
        Exit Sub
    End If
  
  If Me.TextBox2.Value <> "" And CDate(Me.TextBox1.Value) > CDate(Me.TextBox2.Value) Then
   
    MsgBox "Start Date can't greater than End Date.", vbOKOnly + vbInformation, "Start Date"
    Call MyCalendar.DatePicker(Me.TextBox1)
    Exit Sub
    
  End If
  
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
  Exit Sub

ErrorHandler:
    
  MsgBox Err.Description, vbCritical, "Error"
  
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
   
End Sub

Private Sub Cal2_Click()
   
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
   
    
    ''On Error GoTo ErrorHandler
    
    If Me.TextBox1.Value = "" Then
   
        MsgBox "Please select the start date first.", vbOKOnly + vbInformation, "End Date"
    
        Exit Sub
        
    End If
        
    Call MyCalendar.DatePicker(Me.TextBox2)
    
    If CDate(Me.TextBox1.Value) > CDate(Me.TextBox2.Value) Then
   
        MsgBox "End Date can't be after the Start date.", vbOKOnly + vbInformation, "End Date"
        Call MyCalendar.DatePicker(Me.TextBox2)
        Exit Sub
        
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    

End Sub

 
Private Sub CommandButton4_Click()
    
   
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    
    ''On Error GoTo ErrorHandler

    If VBA.IsDate(Me.TextBox1.Value) = False Or VBA.IsDate(Me.TextBox2.Value) = False Then
        MsgBox "Incorrect Date", vbCritical, "Time and Motion"
        Exit Sub
    End If
    
    
    Call FetchRawData(Me.TextBox1, Me.TextBox2)

    Call ExportDataToExcel
    
    Unload Me
    DoEvents
    
    Application.Wait (Now + TimeValue("00:00:01"))
    ThisWorkbook.Activate
    
    frmHome.Show
    DoEvents
        
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    
    MsgBox Err.Description, vbCritical, "Error"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    
End Sub

Private Sub CommandButton5_Click()
    Unload Me
End Sub
 
Private Sub UserForm_Activate()
 
    If Me.TextBox1.Value = "" Then
        Me.TextBox1.Value = VBA.Format(Date - 1, "DD - MMM - YYYY")
    End If
    
    If Me.TextBox2.Value = "" Then
        Me.TextBox2.Value = VBA.Format(Date, "DD - MMM - YYYY")
    End If
 
End Sub

