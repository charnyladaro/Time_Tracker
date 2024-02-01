VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   3924
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5400
   OleObjectBlob   =   "MyCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public InputDate

Private Sub cmbMonth_Change()
    
    Call Change_Heading             'Change Calendar as per Month Selection
    Call HighlightInputDate         ' Highlight day if inptut date is available
        
End Sub

Private Sub cmbYear_Change()
    
    Call Change_Heading             'Change Calendar as per Year Selection
    Call HighlightInputDate         ' Highlight day if inptut date is available
    
End Sub

'Subroutine to Change the Heading of Calendar on Month and Year Change
Private Sub Change_Heading()

    Dim iMonth As Integer
    Dim iYear As Integer
    
    iMonth = Val(Me.cmbMonth.ListIndex) + 1
    iYear = Val(Me.cmbYear.Value)
    
    On Error Resume Next
    
    Me.lblMonthName.Caption = Format(DateSerial(iYear, iMonth, 1), "MMMM YYYY")
    
    Call AssignDay(iMonth, iYear)
    
    On Error GoTo 0

End Sub

Private Sub lblLeft_Click()

    Dim iCurrentMonth As Integer
    Dim iFirstYear As Integer
    Dim iCurrentYear As Integer
    
    iFirstYear = Me.cmbYear.List(0)
    iCurrentYear = Me.cmbYear.List(Me.cmbYear.ListIndex)
    iCurrentMonth = Me.cmbMonth.ListIndex + 1
    
    If iCurrentMonth = 1 Then
        If iCurrentYear > iFirstYear Then
            Me.cmbYear.Value = iCurrentYear - 1
            Me.cmbMonth.Value = Me.cmbMonth.List(11)
        Else
            
            MsgBox "Calendar doesn't support year before " & iCurrentYear & "."
            
        End If
    Else
        
        Me.cmbMonth.Value = Me.cmbMonth.List(Me.cmbMonth.ListIndex - 1)
    
    End If
   
End Sub

Private Sub lblLeft_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Me.lblLeft.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub lblLeft_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.lblLeft.SpecialEffect = fmSpecialEffectFlat

End Sub


Private Sub lblRight_Click()

    Dim iCurrentMonth As Integer
    Dim iLastYear As Integer
    Dim iCurrentYear As Integer
    
    iLastYear = Me.cmbYear.List(Me.cmbYear.ListCount - 1)
    iCurrentYear = Me.cmbYear.List(Me.cmbYear.ListIndex)
    iCurrentMonth = Me.cmbMonth.ListIndex + 1
    
    If iCurrentMonth = 12 Then
        If iCurrentYear < iLastYear Then
            Me.cmbYear.Value = iCurrentYear + 1
            Me.cmbMonth.Value = Me.cmbMonth.List(0)
        Else
        
             MsgBox "Calendar doesn't support year after " & iLastYear & "."
        
        End If
        
    Else
        
        Me.cmbMonth.Value = Me.cmbMonth.List(iCurrentMonth)
    
    End If
    
    
    
End Sub


Private Sub lblRight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Me.lblRight.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub lblRight_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.lblRight.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub UserForm_Activate()
     
    
    If IsDate(InputDate) Then
        
        Dim iDay As Integer
        Dim iMonth As Integer
        Dim sMonthName As String
        Dim iYear As Integer
        Dim iLastYear As Integer
        
        iDay = Val(Format(InputDate, "dd"))
        iMonth = Val(Format(InputDate, "mm"))
        sMonthName = Format(InputDate, "mmmm")
        iYear = Val(Format(InputDate, "YYYY"))
        iLastYear = Me.cmbYear.List(Me.cmbYear.ListCount - 1)
        
        If iYear > iLastYear Then Exit Sub      ' If year of date provided is out of range then do nothing
        
        Me.cmbMonth.Value = sMonthName          ' Assigning Month to Month Combobox
        Me.cmbYear.Value = iYear                ' Assigning Year to Year Combobox
        
        Call Highlight_SelectedDate(iDay)       ' Subroutine to highlight the selected day coming from input parameter
        
        
    End If

    
End Sub

Private Sub UserForm_Initialize()

    Call Add_Month                              ' Subroutine to add month in Month Dropdown
    
    Call Add_Year                               ' Subroutine to add year in Year Dropdown
    
    Me.txtDay.Value = ""
    
    Me.cmbMonth.Value = Format(Date, "MMMM")
    
    Me.cmbYear.Value = Format(Date, "YYYY")
    
    Me.lblMonthName.Caption = Format(Date, "MMMM YYYY")
    
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iMonth = Val(Me.cmbMonth.ListIndex) + 1
    iYear = Val(Me.cmbYear.Value)
    
    On Error Resume Next
    
    Call AssignDay(iMonth, iYear)
    
    On Error GoTo 0

End Sub

Private Sub Add_Month()

    Dim i As Integer
    
    i = 1
    
    Me.cmbMonth.Clear
    
    For i = 1 To 12

        Me.cmbMonth.AddItem Format(DateSerial(1, i, 1), "MMMM")
        
    Next i

End Sub


Private Sub Add_Year()


    Dim i As Integer
    Dim iCurrentYear As Integer
    
    iCurrentYear = Val(Format(Date, "YYYY")) - 1
    
    Me.cmbYear.Clear
    
    For i = iCurrentYear To (iCurrentYear + 50)

        Me.cmbYear.AddItem i
        
    Next i
    

End Sub

Private Sub HighlightInputDate() 'Highlight Day if Input date and available date are matching

    If IsDate(InputDate) Then
        
        Dim iDay As Integer
        Dim sMonthName As String
        Dim iYear As Integer
        
        
        iDay = Val(Format(InputDate, "dd"))
        sMonthName = Format(InputDate, "mmmm")
        iYear = Val(Format(InputDate, "YYYY"))
        
        If iYear = Val(Me.cmbYear.Value) And sMonthName = Me.cmbMonth.Value Then
        
            Call Highlight_SelectedDate(iDay)       ' Subroutine to highlight the selected day coming from input parameter
        Else
        
            Me.Tick.Visible = False
        
        End If
        
    End If


End Sub


Private Sub AssignDay(month As Integer, year As Integer)

    Dim dDate As Date
    
    Dim iLastDay As Integer
    
    Dim iWeek2FirstDay As Integer
    
    Dim iCurrentDay As Integer
    
    Dim sDay As String
    
    If Not IsNumeric(month) Or Not IsNumeric(year) Then Exit Sub
       
    dDate = DateSerial(year, month, 1)
    
    iLastDay = Format(Application.WorksheetFunction.EoMonth(dDate, 0), "dd")
    
    sDay = Format(dDate, "Ddd")

    'Week1
    Me.D1.Caption = IIf(sDay = "Sun", "1", "")                                                      'Sun
    Me.D2.Caption = IIf(Me.D1.Caption <> "", Val(Me.D1.Caption) + 1, IIf(sDay = "Mon", "1", ""))    'Mon
    Me.D3.Caption = IIf(Me.D2.Caption <> "", Val(Me.D2.Caption) + 1, IIf(sDay = "Tue", "1", ""))    'Tue
    Me.D4.Caption = IIf(Me.D3.Caption <> "", Val(Me.D3.Caption) + 1, IIf(sDay = "Wed", "1", ""))    'Wed
    Me.D5.Caption = IIf(Me.D4.Caption <> "", Val(Me.D4.Caption) + 1, IIf(sDay = "Thu", "1", ""))    'Thu
    Me.D6.Caption = IIf(Me.D5.Caption <> "", Val(Me.D5.Caption) + 1, IIf(sDay = "Fri", "1", ""))    'Fri
    Me.D7.Caption = IIf(Me.D6.Caption <> "", Val(Me.D6.Caption) + 1, IIf(sDay = "Sat", "1", ""))    'Sat
    
    iWeek2FirstDay = Val(Me.D7.Caption)
    
    'Week2
    Me.D8.Caption = IIf(iWeek2FirstDay + 1 <= iLastDay, iWeek2FirstDay + 1, "")
    Me.D9.Caption = IIf(iWeek2FirstDay + 2 <= iLastDay, iWeek2FirstDay + 2, "")
    Me.D10.Caption = IIf(iWeek2FirstDay + 3 <= iLastDay, iWeek2FirstDay + 3, "")
    Me.D11.Caption = IIf(iWeek2FirstDay + 4 <= iLastDay, iWeek2FirstDay + 4, "")
    Me.D12.Caption = IIf(iWeek2FirstDay + 5 <= iLastDay, iWeek2FirstDay + 5, "")
    Me.D13.Caption = IIf(iWeek2FirstDay + 6 <= iLastDay, iWeek2FirstDay + 6, "")
    Me.D14.Caption = IIf(iWeek2FirstDay + 7 <= iLastDay, iWeek2FirstDay + 7, "")
    
    'Week3
    Me.D15.Caption = IIf(iWeek2FirstDay + 8 <= iLastDay, iWeek2FirstDay + 8, "")
    Me.D16.Caption = IIf(iWeek2FirstDay + 9 <= iLastDay, iWeek2FirstDay + 9, "")
    Me.D17.Caption = IIf(iWeek2FirstDay + 10 <= iLastDay, iWeek2FirstDay + 10, "")
    Me.D18.Caption = IIf(iWeek2FirstDay + 11 <= iLastDay, iWeek2FirstDay + 11, "")
    Me.D19.Caption = IIf(iWeek2FirstDay + 12 <= iLastDay, iWeek2FirstDay + 12, "")
    Me.D20.Caption = IIf(iWeek2FirstDay + 13 <= iLastDay, iWeek2FirstDay + 13, "")
    Me.D21.Caption = IIf(iWeek2FirstDay + 14 <= iLastDay, iWeek2FirstDay + 14, "")
    
    'Week4
    Me.D22.Caption = IIf(iWeek2FirstDay + 15 <= iLastDay, iWeek2FirstDay + 15, "")
    Me.D23.Caption = IIf(iWeek2FirstDay + 16 <= iLastDay, iWeek2FirstDay + 16, "")
    Me.D24.Caption = IIf(iWeek2FirstDay + 17 <= iLastDay, iWeek2FirstDay + 17, "")
    Me.D25.Caption = IIf(iWeek2FirstDay + 18 <= iLastDay, iWeek2FirstDay + 18, "")
    Me.D26.Caption = IIf(iWeek2FirstDay + 19 <= iLastDay, iWeek2FirstDay + 19, "")
    Me.D27.Caption = IIf(iWeek2FirstDay + 20 <= iLastDay, iWeek2FirstDay + 20, "")
    Me.D28.Caption = IIf(iWeek2FirstDay + 21 <= iLastDay, iWeek2FirstDay + 21, "")
    
    'Week5
    Me.D29.Caption = IIf(iWeek2FirstDay + 22 <= iLastDay, iWeek2FirstDay + 22, "")
    Me.D30.Caption = IIf(iWeek2FirstDay + 23 <= iLastDay, iWeek2FirstDay + 23, "")
    Me.D31.Caption = IIf(iWeek2FirstDay + 24 <= iLastDay, iWeek2FirstDay + 24, "")
    Me.D32.Caption = IIf(iWeek2FirstDay + 25 <= iLastDay, iWeek2FirstDay + 25, "")
    Me.D33.Caption = IIf(iWeek2FirstDay + 26 <= iLastDay, iWeek2FirstDay + 26, "")
    Me.D34.Caption = IIf(iWeek2FirstDay + 27 <= iLastDay, iWeek2FirstDay + 27, "")
    Me.D35.Caption = IIf(iWeek2FirstDay + 28 <= iLastDay, iWeek2FirstDay + 28, "")
    
    'Week6
    Me.D36.Caption = IIf(iWeek2FirstDay + 29 <= iLastDay, iWeek2FirstDay + 29, "")
    Me.D37.Caption = IIf(iWeek2FirstDay + 30 <= iLastDay, iWeek2FirstDay + 30, "")
    Me.D38.Caption = IIf(iWeek2FirstDay + 31 <= iLastDay, iWeek2FirstDay + 31, "")
    Me.D39.Caption = IIf(iWeek2FirstDay + 32 <= iLastDay, iWeek2FirstDay + 32, "")
    Me.D40.Caption = IIf(iWeek2FirstDay + 33 <= iLastDay, iWeek2FirstDay + 33, "")
    Me.D41.Caption = IIf(iWeek2FirstDay + 34 <= iLastDay, iWeek2FirstDay + 34, "")
    Me.D42.Caption = IIf(iWeek2FirstDay + 35 <= iLastDay, iWeek2FirstDay + 35, "")
    
    'Calling Subroutine to change the color
    
    Call Highlight_TodayDate(month, year)
    
    'Disabling Blank Days
    
    Call Disable_Days

End Sub


Private Sub Highlight_TodayDate(month As Integer, year As Integer)
  
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Format(Date, "dd"))
    iMonth = Val(Format(Date, "mm"))
    iYear = Val(Format(Date, "yyyy"))
    
    If iMonth = month And iYear = year Then
        
        Select Case iDay
        
        Case Is = Val(Me.D1.Caption)
                  Me.D1.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D2.Caption)
                  Me.D2.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D3.Caption)
                  Me.D3.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D4.Caption)
                  Me.D4.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D5.Caption)
                  Me.D5.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D6.Caption)
                  Me.D6.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D7.Caption)
                  Me.D7.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D8.Caption)
                  Me.D8.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D9.Caption)
                  Me.D9.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D10.Caption)
                  Me.D10.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D11.Caption)
                  Me.D11.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D12.Caption)
                  Me.D12.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D13.Caption)
                  Me.D13.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D14.Caption)
                  Me.D14.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D15.Caption)
                  Me.D15.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D16.Caption)
                  Me.D16.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D17.Caption)
                  Me.D17.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D18.Caption)
                  Me.D18.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D19.Caption)
                  Me.D19.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D20.Caption)
                  Me.D20.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D21.Caption)
                  Me.D21.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D22.Caption)
                  Me.D22.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D23.Caption)
                  Me.D23.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D24.Caption)
                  Me.D24.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D25.Caption)
                  Me.D25.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D26.Caption)
                  Me.D26.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D27.Caption)
                  Me.D27.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D28.Caption)
                  Me.D28.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D29.Caption)
                  Me.D29.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D30.Caption)
                  Me.D30.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D31.Caption)
                  Me.D31.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D32.Caption)
                  Me.D32.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D33.Caption)
                  Me.D33.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D34.Caption)
                  Me.D34.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D35.Caption)
                  Me.D35.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D36.Caption)
                  Me.D36.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D37.Caption)
                  Me.D37.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D38.Caption)
                  Me.D38.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D39.Caption)
                  Me.D39.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D40.Caption)
                  Me.D40.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D41.Caption)
                  Me.D41.BackColor = &HC0FFC0
                  
        Case Is = Val(Me.D42.Caption)
                  Me.D42.BackColor = &HC0FFC0
        
        End Select
      
    Else
    
        Call DefaultColor
    
    End If
   
End Sub


'Default Color Code

Private Sub DefaultColor()

        Me.D1.BackColor = &H8000000F
        Me.D2.BackColor = &H8000000F
        Me.D3.BackColor = &H8000000F
        Me.D4.BackColor = &H8000000F
        Me.D5.BackColor = &H8000000F
        Me.D6.BackColor = &H8000000F
        Me.D7.BackColor = &H8000000F
        Me.D8.BackColor = &H8000000F
        Me.D9.BackColor = &H8000000F
        Me.D10.BackColor = &H8000000F
        Me.D11.BackColor = &H8000000F
        Me.D12.BackColor = &H8000000F
        Me.D13.BackColor = &H8000000F
        Me.D14.BackColor = &H8000000F
        Me.D15.BackColor = &H8000000F
        Me.D16.BackColor = &H8000000F
        Me.D17.BackColor = &H8000000F
        Me.D18.BackColor = &H8000000F
        Me.D19.BackColor = &H8000000F
        Me.D20.BackColor = &H8000000F
        Me.D21.BackColor = &H8000000F
        Me.D22.BackColor = &H8000000F
        Me.D23.BackColor = &H8000000F
        Me.D24.BackColor = &H8000000F
        Me.D25.BackColor = &H8000000F
        Me.D26.BackColor = &H8000000F
        Me.D27.BackColor = &H8000000F
        Me.D28.BackColor = &H8000000F
        Me.D29.BackColor = &H8000000F
        Me.D30.BackColor = &H8000000F
        Me.D31.BackColor = &H8000000F
        Me.D32.BackColor = &H8000000F
        Me.D33.BackColor = &H8000000F
        Me.D34.BackColor = &H8000000F
        Me.D35.BackColor = &H8000000F
        Me.D36.BackColor = &H8000000F
        Me.D37.BackColor = &H8000000F
        Me.D38.BackColor = &H8000000F
        Me.D39.BackColor = &H8000000F
        Me.D40.BackColor = &H8000000F
        Me.D41.BackColor = &H8000000F
        Me.D42.BackColor = &H8000000F

End Sub

'Subroutine to highlight the selected date coming as input paramtere for Calendar
Private Sub Highlight_SelectedDate(iDay As Integer)
  
        Select Case iDay
        
        Case Is = Val(Me.D1.Caption)
                  Me.D1.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D1.Left + 23
                  Me.Tick.Top = Me.D1.Top
                  Me.Tick.Visible = True

                  
        Case Is = Val(Me.D2.Caption)
                  Me.D2.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D2.Left + 23
                  Me.Tick.Top = Me.D2.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D3.Caption)
                  Me.D3.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D3.Left + 23
                  Me.Tick.Top = Me.D3.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D4.Caption)
                  Me.D4.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D4.Left + 23
                  Me.Tick.Top = Me.D4.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D5.Caption)
                  Me.D5.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D5.Left + 23
                  Me.Tick.Top = Me.D5.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D6.Caption)
                  Me.D6.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D6.Left + 23
                  Me.Tick.Top = Me.D6.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D7.Caption)
                  Me.D7.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D7.Left + 23
                  Me.Tick.Top = Me.D7.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D8.Caption)
                  Me.D8.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D8.Left + 23
                  Me.Tick.Top = Me.D8.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D9.Caption)
                  Me.D9.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D9.Left + 23
                  Me.Tick.Top = Me.D9.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D10.Caption)
                  Me.D10.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D10.Left + 23
                  Me.Tick.Top = Me.D10.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D11.Caption)
                  Me.D11.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D11.Left + 23
                  Me.Tick.Top = Me.D11.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D12.Caption)
                  Me.D12.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D12.Left + 23
                  Me.Tick.Top = Me.D12.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D13.Caption)
                  Me.D13.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D13.Left + 23
                  Me.Tick.Top = Me.D13.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D14.Caption)
                  Me.D14.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D14.Left + 23
                  Me.Tick.Top = Me.D14.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D15.Caption)
                  Me.D15.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D15.Left + 23
                  Me.Tick.Top = Me.D15.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D16.Caption)
                  Me.D16.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D16.Left + 23
                  Me.Tick.Top = Me.D16.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D17.Caption)
                  Me.D17.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D17.Left + 23
                  Me.Tick.Top = Me.D17.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D18.Caption)
                  Me.D18.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D18.Left + 23
                  Me.Tick.Top = Me.D18.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D19.Caption)
                  Me.D19.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D19.Left + 23
                  Me.Tick.Top = Me.D19.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D20.Caption)
                  Me.D20.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D20.Left + 23
                  Me.Tick.Top = Me.D20.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D21.Caption)
                  Me.D21.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D21.Left + 23
                  Me.Tick.Top = Me.D21.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D22.Caption)
                  Me.D22.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D22.Left + 23
                  Me.Tick.Top = Me.D22.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D23.Caption)
                  Me.D23.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D23.Left + 23
                  Me.Tick.Top = Me.D23.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D24.Caption)
                  Me.D24.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D24.Left + 23
                  Me.Tick.Top = Me.D24.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D25.Caption)
                  Me.D25.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D25.Left + 23
                  Me.Tick.Top = Me.D25.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D26.Caption)
                  Me.D26.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D26.Left + 23
                  Me.Tick.Top = Me.D26.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D27.Caption)
                  Me.D27.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D27.Left + 23
                  Me.Tick.Top = Me.D27.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D28.Caption)
                  Me.D28.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D28.Left + 23
                  Me.Tick.Top = Me.D28.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D29.Caption)
                  Me.D29.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D29.Left + 23
                  Me.Tick.Top = Me.D29.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D30.Caption)
                  Me.D30.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D30.Left + 23
                  Me.Tick.Top = Me.D30.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D31.Caption)
                  Me.D31.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D31.Left + 23
                  Me.Tick.Top = Me.D31.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D32.Caption)
                  Me.D32.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D32.Left + 23
                  Me.Tick.Top = Me.D32.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D33.Caption)
                  Me.D33.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D33.Left + 23
                  Me.Tick.Top = Me.D33.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D34.Caption)
                  Me.D34.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D34.Left + 23
                  Me.Tick.Top = Me.D34.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D35.Caption)
                  Me.D35.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D35.Left + 23
                  Me.Tick.Top = Me.D35.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D36.Caption)
                  Me.D36.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D36.Left + 23
                  Me.Tick.Top = Me.D36.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D37.Caption)
                  Me.D37.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D37.Left + 23
                  Me.Tick.Top = Me.D37.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D38.Caption)
                  Me.D38.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D38.Left + 23
                  Me.Tick.Top = Me.D38.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D39.Caption)
                  Me.D39.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D39.Left + 23
                  Me.Tick.Top = Me.D39.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D40.Caption)
                  Me.D40.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D40.Left + 23
                  Me.Tick.Top = Me.D40.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D41.Caption)
                  Me.D41.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D41.Left + 23
                  Me.Tick.Top = Me.D41.Top
                  Me.Tick.Visible = True
                  
        Case Is = Val(Me.D42.Caption)
                  Me.D42.BackColor = &HC0FFFF
                  Me.Tick.Left = Me.D42.Left + 23
                  Me.Tick.Top = Me.D42.Top
                  Me.Tick.Visible = True
        
        End Select
      
   
   
End Sub

'-----------------Lock Non-Value Days-----------------------------

Private Sub Disable_Days()
        
        '--------------------------
        If Me.D1.Caption = "" Then
           Me.D1.Visible = False
        Else
           Me.D1.Visible = True
        End If
        '--------------------------
        If Me.D2.Caption = "" Then
           Me.D2.Visible = False
        Else
           Me.D2.Visible = True
        End If
        '--------------------------
        If Me.D3.Caption = "" Then
           Me.D3.Visible = False
        Else
           Me.D3.Visible = True
        End If
        '--------------------------
        If Me.D4.Caption = "" Then
           Me.D4.Visible = False
        Else
           Me.D4.Visible = True
        End If
        '--------------------------
        If Me.D5.Caption = "" Then
           Me.D5.Visible = False
        Else
           Me.D5.Visible = True
        End If
        '--------------------------
        If Me.D6.Caption = "" Then
           Me.D6.Visible = False
        Else
           Me.D6.Visible = True
        End If
        '--------------------------
        If Me.D7.Caption = "" Then
           Me.D7.Visible = False
        Else
           Me.D7.Visible = True
        End If
        '--------------------------
        If Me.D8.Caption = "" Then
           Me.D8.Visible = False
        Else
           Me.D8.Visible = True
        End If
        '--------------------------
        If Me.D9.Caption = "" Then
           Me.D9.Visible = False
        Else
           Me.D9.Visible = True
        End If
        '--------------------------
        If Me.D10.Caption = "" Then
           Me.D10.Visible = False
        Else
           Me.D10.Visible = True
        End If
        
        '--------------------------
        If Me.D11.Caption = "" Then
           Me.D11.Visible = False
        Else
           Me.D11.Visible = True
        End If
        
        '--------------------------
        If Me.D12.Caption = "" Then
           Me.D12.Visible = False
        Else
           Me.D12.Visible = True
        End If
        
        '--------------------------
        If Me.D13.Caption = "" Then
           Me.D13.Visible = False
        Else
           Me.D13.Visible = True
        End If
        
        '--------------------------
        If Me.D14.Caption = "" Then
           Me.D14.Visible = False
        Else
           Me.D14.Visible = True
        End If
        
        '--------------------------
        If Me.D15.Caption = "" Then
           Me.D15.Visible = False
        Else
           Me.D15.Visible = True
        End If
        
        '--------------------------
        If Me.D16.Caption = "" Then
           Me.D16.Visible = False
        Else
           Me.D16.Visible = True
        End If
        
        '--------------------------
        If Me.D17.Caption = "" Then
           Me.D17.Visible = False
        Else
           Me.D17.Visible = True
        End If
        
        '--------------------------
        If Me.D18.Caption = "" Then
           Me.D18.Visible = False
        Else
           Me.D18.Visible = True
        End If
        
        '--------------------------
        If Me.D19.Caption = "" Then
           Me.D19.Visible = False
        Else
           Me.D19.Visible = True
        End If
   
        '--------------------------
        If Me.D20.Caption = "" Then
           Me.D20.Visible = False
        Else
           Me.D20.Visible = True
        End If
        '--------------------------
        If Me.D21.Caption = "" Then
           Me.D21.Visible = False
        Else
           Me.D21.Visible = True
        End If
        '--------------------------
        If Me.D22.Caption = "" Then
           Me.D22.Visible = False
        Else
           Me.D22.Visible = True
        End If
        '--------------------------
        If Me.D23.Caption = "" Then
           Me.D23.Visible = False
        Else
           Me.D23.Visible = True
        End If
        '--------------------------
        If Me.D24.Caption = "" Then
           Me.D24.Visible = False
        Else
           Me.D24.Visible = True
        End If
        '--------------------------
        If Me.D25.Caption = "" Then
           Me.D25.Visible = False
        Else
           Me.D25.Visible = True
        End If
        '--------------------------
        If Me.D26.Caption = "" Then
           Me.D26.Visible = False
        Else
           Me.D26.Visible = True
        End If
        '--------------------------
        If Me.D27.Caption = "" Then
           Me.D27.Visible = False
        Else
           Me.D27.Visible = True
        End If
        '--------------------------
        If Me.D28.Caption = "" Then
           Me.D28.Visible = False
        Else
           Me.D28.Visible = True
        End If
        '--------------------------
        If Me.D29.Caption = "" Then
           Me.D29.Visible = False
        Else
           Me.D29.Visible = True
        End If
        '--------------------------
        If Me.D30.Caption = "" Then
           Me.D30.Visible = False
        Else
           Me.D30.Visible = True
        End If
        '--------------------------
        If Me.D31.Caption = "" Then
           Me.D31.Visible = False
        Else
           Me.D31.Visible = True
        End If
        '--------------------------
        If Me.D32.Caption = "" Then
           Me.D32.Visible = False
        Else
           Me.D32.Visible = True
        End If
        
        '--------------------------
        If Me.D33.Caption = "" Then
           Me.D33.Visible = False
        Else
           Me.D33.Visible = True
        End If
        '--------------------------
        If Me.D34.Caption = "" Then
           Me.D34.Visible = False
        Else
           Me.D34.Visible = True
        End If
        '--------------------------
        If Me.D35.Caption = "" Then
           Me.D35.Visible = False
        Else
           Me.D35.Visible = True
        End If
        '--------------------------
        If Me.D36.Caption = "" Then
           Me.D36.Visible = False
        Else
           Me.D36.Visible = True
        End If
        '--------------------------
        If Me.D37.Caption = "" Then
           Me.D37.Visible = False
        Else
           Me.D37.Visible = True
        End If
        '--------------------------
        If Me.D38.Caption = "" Then
           Me.D38.Visible = False
        Else
           Me.D38.Visible = True
        End If
        '--------------------------
        If Me.D39.Caption = "" Then
           Me.D39.Visible = False
        Else
           Me.D39.Visible = True
        End If
        '--------------------------
        If Me.D40.Caption = "" Then
           Me.D40.Visible = False
        Else
           Me.D40.Visible = True
        End If
        '--------------------------
        If Me.D41.Caption = "" Then
           Me.D41.Visible = False
        Else
           Me.D41.Visible = True
        End If
        '--------------------------
        If Me.D42.Caption = "" Then
           Me.D42.Visible = False
        Else
           Me.D42.Visible = True
        End If
       
End Sub



'----------------------------------------------------------------------------------------------


'----------------Code to Getting Date from Selection -------------------------------------------

Private Sub D1_Click()
    
    If Me.D1.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D1.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D1.Left + 23
    Me.Tick.Top = Me.D1.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D2_Click()
    
    If Me.D2.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D2.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D2.Left + 23
    Me.Tick.Top = Me.D2.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub


Private Sub D3_Click()
    
    If Me.D3.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D3.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D3.Left + 23
    Me.Tick.Top = Me.D3.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D4_Click()
    
    If Me.D4.Caption = "" Then Exit Sub
    
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D4.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D4.Left + 23
    Me.Tick.Top = Me.D4.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D5_Click()
    
    If Me.D5.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D5.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D5.Left + 23
    Me.Tick.Top = Me.D5.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub


Private Sub D6_Click()
    
    If Me.D6.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D6.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D6.Left + 23
    Me.Tick.Top = Me.D6.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub



Private Sub D7_Click()
    
    If Me.D7.Caption = "" Then Exit Sub
    
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D7.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D7.Left + 23
    Me.Tick.Top = Me.D7.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D8_Click()
    
    If Me.D8.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D8.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D8.Left + 23
    Me.Tick.Top = Me.D8.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D9_Click()
    
    If Me.D9.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D9.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D9.Left + 23
    Me.Tick.Top = Me.D9.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D10_Click()
    
    If Me.D10.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D10.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D10.Left + 23
    Me.Tick.Top = Me.D10.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D11_Click()
    
    If Me.D11.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D11.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D11.Left + 23
    Me.Tick.Top = Me.D11.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub


Private Sub D12_Click()
    
    If Me.D12.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D12.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D12.Left + 23
    Me.Tick.Top = Me.D12.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D13_Click()
    
    If Me.D13.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D13.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D13.Left + 23
    Me.Tick.Top = Me.D13.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D14_Click()
    
    If Me.D14.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D14.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D14.Left + 23
    Me.Tick.Top = Me.D14.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D15_Click()
    
    If Me.D15.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D15.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    Me.Tick.Left = Me.D15.Left + 23
    Me.Tick.Top = Me.D15.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D16_Click()
    
    If Me.D16.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D16.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D16.Left + 23
    Me.Tick.Top = Me.D16.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D17_Click()
    
    If Me.D17.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D17.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D17.Left + 23
    Me.Tick.Top = Me.D17.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D18_Click()
    
    If Me.D18.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D18.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D18.Left + 23
    Me.Tick.Top = Me.D18.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D19_Click()
    
    If Me.D19.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D19.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D19.Left + 23
    Me.Tick.Top = Me.D19.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D20_Click()
    
    If Me.D20.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D20.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D20.Left + 23
    Me.Tick.Top = Me.D20.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D21_Click()
    
    If Me.D21.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D21.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D21.Left + 23
    Me.Tick.Top = Me.D21.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D22_Click()
    
    If Me.D22.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D22.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D22.Left + 23
    Me.Tick.Top = Me.D22.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D23_Click()
    
    If Me.D23.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D23.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D23.Left + 23
    Me.Tick.Top = Me.D23.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D24_Click()
    
    If Me.D24.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D24.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    Me.Tick.Left = Me.D24.Left + 23
    Me.Tick.Top = Me.D24.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D25_Click()
    
    If Me.D25.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D25.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D25.Left + 23
    Me.Tick.Top = Me.D25.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D26_Click()
    
    If Me.D26.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D26.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D26.Left + 23
    Me.Tick.Top = Me.D26.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D27_Click()
    
    If Me.D27.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D27.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D27.Left + 23
    Me.Tick.Top = Me.D27.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D28_Click()
    
    If Me.D28.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D28.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D28.Left + 23
    Me.Tick.Top = Me.D28.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D29_Click()
    
    If Me.D29.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D29.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D29.Left + 23
    Me.Tick.Top = Me.D29.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D30_Click()
    
    If Me.D30.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D30.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D30.Left + 23
    Me.Tick.Top = Me.D30.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D31_Click()
    
    If Me.D31.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D31.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D31.Left + 23
    Me.Tick.Top = Me.D31.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub



Private Sub D32_Click()
    
    If Me.D32.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D32.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D32.Left + 23
    Me.Tick.Top = Me.D32.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D33_Click()
    
    If Me.D33.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D33.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D33.Left + 23
    Me.Tick.Top = Me.D33.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub

Private Sub D34_Click()
    
    If Me.D34.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D34.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D34.Left + 23
    Me.Tick.Top = Me.D34.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D35_Click()
    
    If Me.D35.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D35.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D35.Left + 23
    Me.Tick.Top = Me.D35.Top
    Me.Tick.Visible = True
    
    Unload Me
       
    
End Sub

Private Sub D36_Click()
    
    If Me.D36.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D36.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D36.Left + 23
    Me.Tick.Top = Me.D36.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
    
End Sub

Private Sub D37_Click()
    
    If Me.D37.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D37.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D37.Left + 23
    Me.Tick.Top = Me.D37.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D38_Click()
    
    If Me.D38.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    
    iDay = Val(Me.D38.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D38.Left + 23
    Me.Tick.Top = Me.D38.Top
    Me.Tick.Visible = True
    
    Unload Me
    
End Sub


Private Sub D39_Click()
    
    If Me.D39.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D39.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D39.Left + 23
    Me.Tick.Top = Me.D39.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D40_Click()
    
    If Me.D40.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D40.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D40.Left + 23
    Me.Tick.Top = Me.D40.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D41_Click()
    
    If Me.D41.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D41.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    Me.Tick.Left = Me.D41.Left + 23
    Me.Tick.Top = Me.D41.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

Private Sub D42_Click()
    
    If Me.D42.Caption = "" Then Exit Sub
    
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    
    iDay = Val(Me.D42.Caption)
    iMonth = Val(cmbMonth.ListIndex + 1)
    iYear = Val(Me.cmbYear.Value)
    
    Me.txtDay.Value = VBA.Format(DateSerial(iYear, iMonth, iDay), "DD - MMM - YYYY")
    
    
    Me.Tick.Left = Me.D42.Left + 23
    Me.Tick.Top = Me.D42.Top
    Me.Tick.Visible = True
    
    Unload Me
    
    
End Sub

'------------------Applying SpecialEffect on Click -------------------------------

Private Sub D1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D1.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D2.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D3.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D4.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D5.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D6.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D7.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D8.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D9.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D10_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D10.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D11.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D12.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D13.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D14.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D15_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D15.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D16_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D16.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D17_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D17.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D18_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D18.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D19_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D19.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D20_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D20.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D21.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D22.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D23_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D23.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D24_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D24.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D25_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D25.SpecialEffect = fmSpecialEffectSunken

End Sub



Private Sub D26_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D26.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D27_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D27.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D28_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D28.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D29_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D29.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D30_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D30.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D31_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D31.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D32_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D32.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D33_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D33.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D34_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D34.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D35_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D35.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D36_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D36.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D37_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D37.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D38_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D38.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D39_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D39.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D40_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D40.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D41_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D41.SpecialEffect = fmSpecialEffectSunken

End Sub

Private Sub D42_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D42.SpecialEffect = fmSpecialEffectSunken

End Sub

'------------------Applying SpecialEffect on MouseUp -------------------------------

Private Sub D1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D1.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D2.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D3.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D4.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D5.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D6_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D6.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D7_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D7.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D8_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D8.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D9.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D10_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D10.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D11.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D12.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D13.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D14_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D14.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D15_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D15.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D16_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D16.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D17_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D17.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D18_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D18.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D19_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D19.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D20_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D20.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D21_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D21.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D22_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D22.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D23_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D23.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D24_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D24.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D25_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D25.SpecialEffect = fmSpecialEffectEtched

End Sub



Private Sub D26_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D26.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D27_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D27.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D28_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D28.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D29_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D29.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D30_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D30.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D31_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D31.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D32_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D32.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D33_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D33.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D34_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D34.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D35_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D35.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D36_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D36.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D37_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D37.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D38_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D38.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D39_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D39.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D40_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D40.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D41_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D41.SpecialEffect = fmSpecialEffectEtched

End Sub

Private Sub D42_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Me.D42.SpecialEffect = fmSpecialEffectEtched

End Sub



Public Function DatePicker(Optional ByRef DateInput As Object) As String
    
    InputDate = "" ' Clearnig the previous Value of variable
    
    'Check the calendar type and date
    
    'MsgBox (TypeName(DateInput))
    
    If (TypeName(DateInput)) = "TextBox" Then
    
        If IsDate(DateInput.Value) Then
        
            InputDate = CDate(DateInput.Value)
        Else
            InputDate = ""
        End If
        
    ElseIf (TypeName(DateInput)) = "CommandButton" Or (TypeName(DateInput)) = "Label" Then
    
        If IsDate(DateInput.Caption) Then
        
            InputDate = CDate(DateInput.Caption)
        Else
            InputDate = ""
            
        End If
    
    Else
    
        InputDate = ""
        
    End If
    
    Me.Show                 ' Show Calendar Form

    'Assign or return the selected date
    
    
    If (TypeName(DateInput)) = "TextBox" Then
    
        DateInput.Value = Me.txtDay.Value
        
    ElseIf (TypeName(DateInput)) = "CommandButton" Or (TypeName(DateInput)) = "Label" Then
        
        DateInput.Caption = Me.txtDay.Value
       
    Else
    
        DatePicker = Me.txtDay.Value
        
    End If
    
    
    
End Function


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If Me.txtDay.Value = "" And InputDate <> "" Then
    
        Me.txtDay.Value = Format(InputDate, "DD - MMM - YYYY")
        
    End If
    
    
End Sub
