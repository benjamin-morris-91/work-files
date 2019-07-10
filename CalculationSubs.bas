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

Sub SplitSSN()
'Splits the ssn into single digits in the database tab
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

Sub Easy()
'    ActiveWorkbook.Names.Add Name:="Prop21LegalB", RefersToR1C1:="=DropdownInfo!R68C15"
'    ActiveWorkbook.Names.Add Name:="Prop22LegalB", RefersToR1C1:="=DropdownInfo!R69C15"
'    ActiveWorkbook.Names.Add Name:="Prop23LegalB", RefersToR1C1:="=DropdownInfo!R70C15"
'    ActiveWorkbook.Names.Add Name:="Prop24LegalB", RefersToR1C1:="=DropdownInfo!R71C15"
'    ActiveWorkbook.Names.Add Name:="Prop25LegalB", RefersToR1C1:="=DropdownInfo!R72C15"


'    ActiveWorkbook.Names.Add Name:="UniqueEntity1", RefersToR1C1:="=Driver!R1C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity2", RefersToR1C1:="=Driver!R2C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity3", RefersToR1C1:="=Driver!R3C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity4", RefersToR1C1:="=Driver!R4C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity5", RefersToR1C1:="=Driver!R5C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity6", RefersToR1C1:="=Driver!R6C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity7", RefersToR1C1:="=Driver!R7C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity8", RefersToR1C1:="=Driver!R8C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity9", RefersToR1C1:="=Driver!R9C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity10", RefersToR1C1:="=Driver!R10C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity11", RefersToR1C1:="=Driver!R11C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity12", RefersToR1C1:="=Driver!R12C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity13", RefersToR1C1:="=Driver!R13C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity14", RefersToR1C1:="=Driver!R14C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity15", RefersToR1C1:="=Driver!R15C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity16", RefersToR1C1:="=Driver!R16C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity17", RefersToR1C1:="=Driver!R17C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity18", RefersToR1C1:="=Driver!R18C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity19", RefersToR1C1:="=Driver!R19C5"
'    ActiveWorkbook.Names.Add Name:="UniqueEntity20", RefersToR1C1:="=Driver!R20C5"
'
'    ActiveWorkbook.Names.Add Name:="UniqueAmount1", RefersToR1C1:="=Driver!R1C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount2", RefersToR1C1:="=Driver!R2C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount3", RefersToR1C1:="=Driver!R3C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount4", RefersToR1C1:="=Driver!R4C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount5", RefersToR1C1:="=Driver!R5C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount6", RefersToR1C1:="=Driver!R6C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount7", RefersToR1C1:="=Driver!R7C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount8", RefersToR1C1:="=Driver!R8C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount9", RefersToR1C1:="=Driver!R9C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount10", RefersToR1C1:="=Driver!R10C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount11", RefersToR1C1:="=Driver!R11C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount12", RefersToR1C1:="=Driver!R12C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount13", RefersToR1C1:="=Driver!R13C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount14", RefersToR1C1:="=Driver!R14C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount15", RefersToR1C1:="=Driver!R15C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount16", RefersToR1C1:="=Driver!R16C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount17", RefersToR1C1:="=Driver!R17C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount18", RefersToR1C1:="=Driver!R18C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount19", RefersToR1C1:="=Driver!R19C6"
'    ActiveWorkbook.Names.Add Name:="UniqueAmount20", RefersToR1C1:="=Driver!R20C6"

End Sub

Sub FKACheck() 'checks, after saving, if there is a FKA. Then it combines the name w/ the FKA and overrides
'the name of the borrower. This allows it to send to the docs, but not to the database.

Range("CombinedB1FKA") = 0
Range("CombinedB2FKA") = 0
Range("CombinedB3FKA") = 0

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


End Sub

Sub Easy1()

'ActiveWorkbook.Names.Add Name:="UniqueEntity1", RefersToR1C1:="=Driver!R1C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity2", RefersToR1C1:="=Driver!R2C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity3", RefersToR1C1:="=Driver!R3C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity4", RefersToR1C1:="=Driver!R4C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity5", RefersToR1C1:="=Driver!R5C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity6", RefersToR1C1:="=Driver!R6C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity7", RefersToR1C1:="=Driver!R7C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity8", RefersToR1C1:="=Driver!R8C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity9", RefersToR1C1:="=Driver!R9C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity10", RefersToR1C1:="=Driver!R10C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity11", RefersToR1C1:="=Driver!R11C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity12", RefersToR1C1:="=Driver!R12C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity13", RefersToR1C1:="=Driver!R13C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity14", RefersToR1C1:="=Driver!R14C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity15", RefersToR1C1:="=Driver!R15C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity16", RefersToR1C1:="=Driver!R16C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity17", RefersToR1C1:="=Driver!R17C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity18", RefersToR1C1:="=Driver!R18C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity19", RefersToR1C1:="=Driver!R19C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity20", RefersToR1C1:="=Driver!R20C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity21", RefersToR1C1:="=Driver!R21C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity22", RefersToR1C1:="=Driver!R22C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity23", RefersToR1C1:="=Driver!R23C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity24", RefersToR1C1:="=Driver!R24C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity25", RefersToR1C1:="=Driver!R25C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity26", RefersToR1C1:="=Driver!R26C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity27", RefersToR1C1:="=Driver!R27C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity28", RefersToR1C1:="=Driver!R28C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity29", RefersToR1C1:="=Driver!R29C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity30", RefersToR1C1:="=Driver!R30C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity31", RefersToR1C1:="=Driver!R31C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity32", RefersToR1C1:="=Driver!R32C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity33", RefersToR1C1:="=Driver!R33C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity34", RefersToR1C1:="=Driver!R34C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity35", RefersToR1C1:="=Driver!R35C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity36", RefersToR1C1:="=Driver!R36C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity37", RefersToR1C1:="=Driver!R37C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity38", RefersToR1C1:="=Driver!R38C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity39", RefersToR1C1:="=Driver!R39C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity40", RefersToR1C1:="=Driver!R40C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity41", RefersToR1C1:="=Driver!R41C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity42", RefersToR1C1:="=Driver!R42C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity43", RefersToR1C1:="=Driver!R43C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity44", RefersToR1C1:="=Driver!R44C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity45", RefersToR1C1:="=Driver!R45C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity46", RefersToR1C1:="=Driver!R46C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity47", RefersToR1C1:="=Driver!R47C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity48", RefersToR1C1:="=Driver!R48C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity49", RefersToR1C1:="=Driver!R49C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity50", RefersToR1C1:="=Driver!R50C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity51", RefersToR1C1:="=Driver!R51C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity52", RefersToR1C1:="=Driver!R52C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity53", RefersToR1C1:="=Driver!R53C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity54", RefersToR1C1:="=Driver!R54C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity55", RefersToR1C1:="=Driver!R55C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity56", RefersToR1C1:="=Driver!R56C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity57", RefersToR1C1:="=Driver!R57C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity58", RefersToR1C1:="=Driver!R58C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity59", RefersToR1C1:="=Driver!R59C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity60", RefersToR1C1:="=Driver!R60C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity61", RefersToR1C1:="=Driver!R61C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity62", RefersToR1C1:="=Driver!R62C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity63", RefersToR1C1:="=Driver!R63C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity64", RefersToR1C1:="=Driver!R64C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity65", RefersToR1C1:="=Driver!R65C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity66", RefersToR1C1:="=Driver!R66C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity67", RefersToR1C1:="=Driver!R67C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity68", RefersToR1C1:="=Driver!R68C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity69", RefersToR1C1:="=Driver!R69C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity70", RefersToR1C1:="=Driver!R70C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity71", RefersToR1C1:="=Driver!R71C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity72", RefersToR1C1:="=Driver!R72C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity73", RefersToR1C1:="=Driver!R73C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity74", RefersToR1C1:="=Driver!R74C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity75", RefersToR1C1:="=Driver!R75C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity76", RefersToR1C1:="=Driver!R76C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity77", RefersToR1C1:="=Driver!R77C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity78", RefersToR1C1:="=Driver!R78C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity79", RefersToR1C1:="=Driver!R79C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity80", RefersToR1C1:="=Driver!R80C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity81", RefersToR1C1:="=Driver!R81C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity82", RefersToR1C1:="=Driver!R82C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity83", RefersToR1C1:="=Driver!R83C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity84", RefersToR1C1:="=Driver!R84C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity85", RefersToR1C1:="=Driver!R85C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity86", RefersToR1C1:="=Driver!R86C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity87", RefersToR1C1:="=Driver!R87C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity88", RefersToR1C1:="=Driver!R88C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity89", RefersToR1C1:="=Driver!R89C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity90", RefersToR1C1:="=Driver!R90C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity91", RefersToR1C1:="=Driver!R91C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity92", RefersToR1C1:="=Driver!R92C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity93", RefersToR1C1:="=Driver!R93C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity94", RefersToR1C1:="=Driver!R94C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity95", RefersToR1C1:="=Driver!R95C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity96", RefersToR1C1:="=Driver!R96C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity97", RefersToR1C1:="=Driver!R97C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity98", RefersToR1C1:="=Driver!R98C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity99", RefersToR1C1:="=Driver!R99C5"
'ActiveWorkbook.Names.Add Name:="UniqueEntity100", RefersToR1C1:="=Driver!R100C5"

'ActiveWorkbook.Names.Add Name:="UniqueAmount1", RefersToR1C1:="=Driver!R1C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount2", RefersToR1C1:="=Driver!R2C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount3", RefersToR1C1:="=Driver!R3C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount4", RefersToR1C1:="=Driver!R4C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount5", RefersToR1C1:="=Driver!R5C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount6", RefersToR1C1:="=Driver!R6C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount7", RefersToR1C1:="=Driver!R7C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount8", RefersToR1C1:="=Driver!R8C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount9", RefersToR1C1:="=Driver!R9C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount10", RefersToR1C1:="=Driver!R10C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount11", RefersToR1C1:="=Driver!R11C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount12", RefersToR1C1:="=Driver!R12C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount13", RefersToR1C1:="=Driver!R13C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount14", RefersToR1C1:="=Driver!R14C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount15", RefersToR1C1:="=Driver!R15C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount16", RefersToR1C1:="=Driver!R16C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount17", RefersToR1C1:="=Driver!R17C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount18", RefersToR1C1:="=Driver!R18C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount19", RefersToR1C1:="=Driver!R19C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount20", RefersToR1C1:="=Driver!R20C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount21", RefersToR1C1:="=Driver!R21C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount22", RefersToR1C1:="=Driver!R22C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount23", RefersToR1C1:="=Driver!R23C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount24", RefersToR1C1:="=Driver!R24C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount25", RefersToR1C1:="=Driver!R25C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount26", RefersToR1C1:="=Driver!R26C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount27", RefersToR1C1:="=Driver!R27C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount28", RefersToR1C1:="=Driver!R28C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount29", RefersToR1C1:="=Driver!R29C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount30", RefersToR1C1:="=Driver!R30C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount31", RefersToR1C1:="=Driver!R31C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount32", RefersToR1C1:="=Driver!R32C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount33", RefersToR1C1:="=Driver!R33C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount34", RefersToR1C1:="=Driver!R34C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount35", RefersToR1C1:="=Driver!R35C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount36", RefersToR1C1:="=Driver!R36C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount37", RefersToR1C1:="=Driver!R37C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount38", RefersToR1C1:="=Driver!R38C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount39", RefersToR1C1:="=Driver!R39C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount40", RefersToR1C1:="=Driver!R40C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount41", RefersToR1C1:="=Driver!R41C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount42", RefersToR1C1:="=Driver!R42C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount43", RefersToR1C1:="=Driver!R43C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount44", RefersToR1C1:="=Driver!R44C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount45", RefersToR1C1:="=Driver!R45C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount46", RefersToR1C1:="=Driver!R46C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount47", RefersToR1C1:="=Driver!R47C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount48", RefersToR1C1:="=Driver!R48C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount49", RefersToR1C1:="=Driver!R49C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount50", RefersToR1C1:="=Driver!R50C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount51", RefersToR1C1:="=Driver!R51C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount52", RefersToR1C1:="=Driver!R52C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount53", RefersToR1C1:="=Driver!R53C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount54", RefersToR1C1:="=Driver!R54C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount55", RefersToR1C1:="=Driver!R55C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount56", RefersToR1C1:="=Driver!R56C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount57", RefersToR1C1:="=Driver!R57C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount58", RefersToR1C1:="=Driver!R58C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount59", RefersToR1C1:="=Driver!R59C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount60", RefersToR1C1:="=Driver!R60C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount61", RefersToR1C1:="=Driver!R61C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount62", RefersToR1C1:="=Driver!R62C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount63", RefersToR1C1:="=Driver!R63C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount64", RefersToR1C1:="=Driver!R64C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount65", RefersToR1C1:="=Driver!R65C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount66", RefersToR1C1:="=Driver!R66C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount67", RefersToR1C1:="=Driver!R67C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount68", RefersToR1C1:="=Driver!R68C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount69", RefersToR1C1:="=Driver!R69C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount70", RefersToR1C1:="=Driver!R70C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount71", RefersToR1C1:="=Driver!R71C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount72", RefersToR1C1:="=Driver!R72C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount73", RefersToR1C1:="=Driver!R73C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount74", RefersToR1C1:="=Driver!R74C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount75", RefersToR1C1:="=Driver!R75C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount76", RefersToR1C1:="=Driver!R76C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount77", RefersToR1C1:="=Driver!R77C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount78", RefersToR1C1:="=Driver!R78C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount79", RefersToR1C1:="=Driver!R79C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount80", RefersToR1C1:="=Driver!R80C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount81", RefersToR1C1:="=Driver!R81C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount82", RefersToR1C1:="=Driver!R82C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount83", RefersToR1C1:="=Driver!R83C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount84", RefersToR1C1:="=Driver!R84C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount85", RefersToR1C1:="=Driver!R85C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount86", RefersToR1C1:="=Driver!R86C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount87", RefersToR1C1:="=Driver!R87C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount88", RefersToR1C1:="=Driver!R88C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount89", RefersToR1C1:="=Driver!R89C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount90", RefersToR1C1:="=Driver!R90C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount91", RefersToR1C1:="=Driver!R91C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount92", RefersToR1C1:="=Driver!R92C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount93", RefersToR1C1:="=Driver!R93C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount94", RefersToR1C1:="=Driver!R94C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount95", RefersToR1C1:="=Driver!R95C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount96", RefersToR1C1:="=Driver!R96C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount97", RefersToR1C1:="=Driver!R97C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount98", RefersToR1C1:="=Driver!R98C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount99", RefersToR1C1:="=Driver!R99C6"
'ActiveWorkbook.Names.Add Name:="UniqueAmount100", RefersToR1C1:="=Driver!R100C6"

End Sub
