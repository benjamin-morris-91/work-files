Attribute VB_Name = "NewLoan"
Option Explicit

Sub LoadNewLoan() ' Called when user clicks the New Loan button.
    
    Application.Calculation = xlCalculationManual
    
If MsgBox("Do you want to clear the form for a new loan?", vbYesNo) = vbNo Then
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
Else
    Call ClearAll
    Call calculateTotalAmountDue
    Call autofill 'Fills loan presets, calls calcProcessingFee
    Call prefillPropAns
    Call prefillBorrowerInfo
End If
    
    Application.Calculation = xlCalculationAutomatic

End Sub
