VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Load_Loan 
   Caption         =   "Load Loan"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "Load_Loan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Load_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButton1_Click()


Application.ScreenUpdating = False
Range("LNToLoad") = Load_Loan.LoanToLoad

'**************************
'Don't think I need the below code and the TypeSelected named range in excel.
'**************************

'If Range("TypeSelected") = "Top" Then
'    Range("LNToLoad") = Load_Loan.LoanToLoad
'Else
'    Range("LNToLoad") = Load_Loan.NameToLoad
'End If

Unload Load_Loan

'Call LoadFromDatabase Already does this before now?
Call LoadLoan

Application.ScreenUpdating = True


End Sub

Private Sub CommandButton2_Click()

Unload Load_Loan
SearchBox.Show

End Sub



'Private Sub LoanToLoad_Change()
'
'Range("TypeSelected") = "Top"
'
'End Sub
'
'Private Sub NameToLoad_Change()
'
'Range("TypeSelected") = "Bottom"
'
'End Sub

