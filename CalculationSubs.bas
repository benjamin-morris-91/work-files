Attribute VB_Name = "CalculationSubs"
Sub SolvePayment()

'Read current settings
CurrentCalcMode = Application.Calculation
CurrentMaxIterations = Application.MaxIterations
CurrentIterations = Application.Iteration
CurrentMaxChange = Application.MaxChange

'Adjust settings for GoalSeek
Application.Iteration = True
Application.MaxIterations = 1
CurrentMaxChange = 0.005
Application.Calculation = xlCalculationAutomatic

'Run goal settings
Worksheets("New Amort").Range("R9").GoalSeek Goal:=0, ChangingCell:=Worksheets("New Amort").Range("D9")
'Sheets("APR Calc").Select
'Range("D9").Value = Range("'New Amort'!D9").Value
'    Range("S9").GoalSeek Goal:=0, ChangingCell:=Range("D9")

If Worksheets("New Loan").Range("H9").Value < Worksheets("New Loan").Range("H10").Value Then
    Worksheets("New Amort").Range("D9").Value = Worksheets("New Amort").Range("D9").Value + 0.01
End If

'Restore Defaults
Application.Iteration = CurrentIterations
Application.MaxChange = CurrentMaxChange
Application.Calculation = CurrentCalcMode
Application.MaxIterations = CurrentMaxIterations

'Making changes to match law office loan docs. Leaving stubbed out code in case of reversal
Range("LastPaymentDate") = Worksheets("New Loan").Range("K10")
Range("APR") = Worksheets("New Loan").Range("D27")
Range("AmountFinanced") = Worksheets("New Loan").Range("F27")
'Range("MonthlyPayment") = Worksheets("New Loan").Range("H9")
'Range("FinalPayment") = Worksheets("New Loan").Range("H10")
'Range("FinanceCharge") = Worksheets("New Loan").Range("H27")
'Range("TotalOfPayments") = Worksheets("New Loan").Range("J27")

'Pulled from the 30-360 tab
'Range("MonthlyPayment") = Worksheets("New Loan").Range("R14")
'Range("FinalPayment") = Worksheets("New Loan").Range("R15")
'Range("TotalOfPayments") = Worksheets("New Loan").Range("R16")
'Range("FinanceCharge") = Worksheets("New Loan").Range("R17")

'Pulled from the Deferred to Maturity 30-360 tab
Range("MonthlyPayment") = Worksheets("New Loan").Range("R19")
Range("FinalPayment") = Worksheets("New Loan").Range("R20")
Range("TotalOfPayments") = Worksheets("New Loan").Range("R21")
Range("FinanceCharge") = Worksheets("New Loan").Range("R22")


End Sub

Sub updatePropertyAmountDue()

Dim i As Integer
Dim j As Integer
Dim nextTotal As String
Dim nextProp As String
Dim totalProp As Single
Dim totalTax As Single

For i = 1 To Range("NumberofProperties")
    nextTotal = "Prop" & i & "TotalAmountDue"
    For j = 1 To 4  'Cycles through 1-4 potential amounts in the PropiAmountDuej, adds them to totalProp
        nextProp = "Prop" & i & "AmountDue" & j
        totalProp = totalProp + Range(nextProp).Value
    Next j
    Range(nextTotal) = totalProp
    totalTax = totalTax + totalProp 'Updates the totalTax variable to keep a running tab for the TotalTaxAmounts field
    totalProp = 0 'resets totalProp so next property doesn't include it.
Next i

End Sub

Sub UniqueLoans() 'Assigns the number of unique loan account numbers to the named range UniqueLoanNumbers

Range("UniqueLoanNumbers").Value2 = Range("UniqueCalculator")

End Sub

Sub AddDupes() 'Figures out the

Dim totalNames As Range
Dim totalSums As Range
Dim numOfUniqueItems As Integer
Dim numOfUniqueSworns As Integer
Dim i As Integer

Worksheets("Driver").Activate
Application.ScreenUpdating = False

Set totalNames = ActiveSheet.Range("A1", Range("A1").End(xlDown))
Set totalSums = ActiveSheet.Range("B1", Range("B1").End(xlDown))
    
    Range("R101").Value = 0
    totalNames.Copy
    Range("C1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("R1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ActiveSheet.Range("R1", Range("R1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo 'Remove duplicates
    
    'REMOVES all zeros, not just the cells only containing zeros
    
    'Remove zero entry
    Columns("R").Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("C").Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
       
    Columns("E").ClearContents 'To clear potential old values
    Columns("F").ClearContents 'To clear potential old values
    
    'Shift Colume C up so that no blank cells are between elements
    Range("R1:R101").SpecialCells(xlCellTypeBlanks).Delete ' Need to delete values, not formulas
    Range("R1:R100").Copy
    Range("E1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'determines the number of unique cells in Column C
    On Error Resume Next
    numOfUniqueItems = Range("R:R").Cells.SpecialCells(xlCellTypeConstants).Count
    If Err.number <> 0 Then
        Range("numOfUniqueItems").Value = 0
    End If
    
    Range("NumberOfUniqueEntities").Value = numOfUniqueItems

    'Update the number of sworn documents needed
    numOfUniqueSworns = Range("C:C").Cells.SpecialCells(xlCellTypeConstants).Count
    Range("NumberOfSworns").Value = numOfUniqueSworns
    
    'Starting the SUMIFS
    For i = 1 To numOfUniqueItems
        Range("F" & i) = Application.SumIfs(totalSums, totalNames, Range("E" & i))
    Next i
    
    Worksheets("Driver").Range("F1:F20").Select
    With Selection
        .NumberFormat = "$#,##0.00"
        .Value = .Value
    End With

Application.ScreenUpdating = True
Application.CutCopyMode = False
Worksheets("Driver").Range("A1").Select


End Sub

Sub SplitSSN() 'Splits the ssn into single digits in the database tab
'MID(TEXT(F13,"000000000"),1,1) The first "1" increases until 9
Dim str As String
Dim range2 As String
Dim i As Integer
Dim B1 As String

For i = 1 To 9 'For Borrower1
    range2 = "B1SSN0" & i
    str = Mid(Format(Worksheets("Database").Range("F13"), "000000000"), i, 1)
    Range(range2) = str
Next i

For i = 1 To 9 'For Borrower2
    range2 = "B2SSN0" & i
    str = Mid(Format(Worksheets("Database").Range("F14"), "000000000"), i, 1)
    Range(range2) = str
Next i

For i = 1 To 9 'For Borrower3
    range2 = "B3SSN0" & i
    str = Mid(Format(Worksheets("Database").Range("F15"), "000000000"), i, 1)
    Range(range2) = str
Next i

End Sub

Sub FKACheck() 'checks, after saving, if there is a FKA. Then it combines the name w/ the FKA and overrides
'the name of the borrower. This allows it to send to the docs, but not to the database.

'******************************************
'Updated 7/12/19 - BM to allow individual names to flow through to Affadavit of Identity
'Added the Range(Borrower1-3NameStorage) to the IF statements and the declaration of Range(B1-3NameStorage)
'******************************************

Range("CombinedB1FKA") = 0
Range("CombinedB2FKA") = 0
Range("CombinedB3FKA") = 0
Range("Borrower1NameStorage") = Range("Borrower1Name")
Range("Borrower2NameStorage") = Range("Borrower2Name")
Range("Borrower3NameStorage") = Range("Borrower3Name")

If Range("Borrower1FKA") <> 0 Then
    Call AddB1FKA
    Range("TempFKA1") = Range("Borrower1Name")
    Range("Borrower1Name") = Range("CombinedB1FKA")
End If
If Range("Borrower2FKA") <> 0 Then
    Range("TempFKA2") = Range("Borrower2Name")
    Call AddB2FKA
    Range("Borrower2Name") = Range("CombinedB2FKA")
End If
If Range("Borrower3FKA") <> 0 Then
    Range("TempFKA3") = Range("Borrower3Name")
    Call AddB3FKA
    Range("Borrower3Name") = Range("CombinedB3FKA")
End If

End Sub

Sub FKASwap()

If Range("TempFKA1") <> 0 Then
    Range("Borrower1Name") = Range("TempFKA1")
    Range("TempFKA1") = ""
End If

If Range("TempFKA2") <> 0 Then
    Range("Borrower2Name") = Range("TempFKA2")
    Range("TempFKA2") = ""
End If

If Range("TempFKA3") <> 0 Then
    Range("Borrower3Name") = Range("TempFKA3")
    Range("TempFKA3") = ""
End If

'Added below lines to counter the affadavit of identity problem with names.
Range("Borrower1NameStorage") = 0
Range("Borrower2NameStorage") = 0
Range("Borrower3NameStorage") = 0

End Sub

Sub ShowPayment()

Dim flag1 As Integer

Call calcProcessingFee

If Range("InterestRate") = 0 Then
    MsgBox "Enter an interest rate before calculating the monthly payment"
ElseIf Range("Term") = 0 Then
    MsgBox "Enter a term before calculating the monthly payment"
ElseIf Range("SigningDate") = 0 Then
    MsgBox "Enter a Target Closing date before calculating the monthly payment"
ElseIf Range("FirstPaymentDate") = 0 Then
    MsgBox "Enter a First Payment Date before calculating the monthly payment"
ElseIf Range("Prop1AmountDue1") = 0 Then
    MsgBox "There needs to be at least one taxing entity entered before calculating the monthly payment"
Else 'run and display the payment
    'MsgBox "The first month's payment is: " & Worksheets("New Loan").Range("R19")
    Worksheets("Sheet1").Range("CG17").Value = Worksheets("New Loan").Range("R19")
End If

End Sub

