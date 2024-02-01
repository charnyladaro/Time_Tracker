Attribute VB_Name = "mdlTodayDate"
Option Explicit

'Code to get the actual login date if system date and time gettting changed after 12 AM in night

Public Function TodayDate() As Variant

'    If [NOW()-TODAY()] >= #12:00:01 AM# And [NOW()-TODAY()] <= #7:00:00 AM# Then
'
'        TodayDate = VBA.Format(Date - 1, "DD MMMM YYYY")
'
'    Else
'
'        TodayDate = VBA.Format(Date, "DD MMMM YYYY")
'
'    End If
    
    TodayDate = VBA.Format(Date, "DD MMMM YYYY")

End Function
