VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatabase 
   Caption         =   "Database Path"
   ClientHeight    =   2412
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7848
   OleObjectBlob   =   "frmDatabase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAssign_Click()
    
    Dim sPath As String
    
    sPath = Me.txtPath.Value
    
    Me.txtPath.BackColor = vbWhite
    
    If Dir(sPath) = "" Then
    
        MsgBox "Please provide the complete valid path along with file name 'Database.mdb'.", vbOKOnly + vbCritical, "Error"
        Me.txtPath.SetFocus
        Me.txtPath.BackColor = vbRed
        Exit Sub
        
    Else
    
        ThisWorkbook.Sheets("Database Path").Range("A2").Value = sPath
        
        ThisWorkbook.Save
        
        Unload Me
        
    End If
    
    MsgBox "Database path successfully set! Restart this file.", vbOKOnly + vbInformation, "Database Path"
    
End Sub

Private Sub cmdDefault_Click()
    
    Dim msgValue As VbMsgBoxResult
    
    msgValue = MsgBox("Do you want to assign the Default database path?", vbYesNo + vbQuestion, "Default")
    
    If msgValue = vbNo Then Exit Sub
    
    Me.txtPath.Value = "Default"
    
    ThisWorkbook.Sheets("Database Path").Range("A2").Value = "Default"
    
    Unload Me
    
    MsgBox "Default path has been updated. Restart this file to see the impact.", vbOKOnly + vbInformation, "Default Path"
    
    
End Sub

