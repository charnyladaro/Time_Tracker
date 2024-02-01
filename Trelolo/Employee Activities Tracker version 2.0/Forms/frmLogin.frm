VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login to Employee Activities Tracker v. 2.0"
   ClientHeight    =   3372
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7452
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogin_Click()

    Call Login
    
End Sub

Private Sub cmdClear_Click()

    Me.txtUserID.Value = ""
    Me.txtPassword.Value = ""
    Me.txtUserID.SetFocus
    
End Sub


Private Sub txtUserID_Change()

End Sub

Private Sub UserForm_Activate()
    
    Me.txtUserID.Value = ""
    Me.txtPassword.Value = ""
    Me.txtUserID.SetFocus
     
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    'Reset the Login and other support details
    
    ThisWorkbook.Sheets("Login Details").Range("A2:D" & Application.Rows.Count).Value = ""
    ThisWorkbook.Sheets("RawData").Range("A2:J" & Application.Rows.Count).Value = ""
    ThisWorkbook.Sheets("UM_Support").Range("A2:F" & Application.Rows.Count).Value = ""
    
End Sub
