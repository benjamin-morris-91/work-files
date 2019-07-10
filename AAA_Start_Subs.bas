Attribute VB_Name = "AAA_Start_Subs"
Option Explicit

Sub Auto_Open() 'Automatically runs when excel opens

Worksheets("Sheet1").Protect Password:="123", UserInterfaceOnly:=True
Call UserFilePath
Call assignFileNames

End Sub
Sub Auto_Close() 'Automatically runs when excel closes

    'Call ClearAll
    'Range("Entity").Select
    
End Sub

Sub UpdateTotals()  ' All updating values should go here. This will run before saving.

Dim closingDate As Date
Dim proposedRescindDate As Date
Dim netWorkDayResult As Integer
Dim daysToAdd As Integer
Dim sundayCheck As Variant
Dim companyHolidays1 As Range

'Updates the totalOtherFees range before saving.
Range("TotalOtherFees") = Range("CourtCost") + Range("ClosingCosts")

'Updates the Prop(1-25)TotalAmountDue fields
Call calculateTotalAmountDue

'Updates the Principal Payment Amount, AmountToTaxCollector, and OtherFeesCharged
Range("AmountToTaxCollector") = Range("TotalTaxAmount") + Range("TotalCourtCosts")
Range("OtherFeesCharged") = Range("ProcessingFee") + Range("UnderwritingFee") + Range("FloodFee") + _
    Range("AssessmentOfValue") + Range("BankruptcySearch") + Range("InternalTitleSearch") + Range("InternalTitleReview") + _
    Range("DocumentPreparationFee") + Range("TitleCurativeFee") + Range("NotaryFee") + Range("MailingFee") + _
    Range("ExternalTitleSearch") + Range("ExternalTitleReview") + Range("RecordingFee")
Range("PrincipalPaymentAmount") = Range("AmountToTaxCollector") + Range("OtherFeesCharged") - Range("CreditAmount")

'Updated AmountFinanced
Range("AmountFinanced") = Range("AmountToTaxCollector") + Range("DocumentPreparationFee") + Range("TitleCurativeFee") + _
    Range("NotaryFee") + Range("MailingFee") + Range("ExternalTitleSearch") + Range("ExternalTitleReview") + Range("RecordingFee")

'Updates the LNPlusName Range
Range("LNPlusName") = Range("LoanNumber") & " - " & Range("Borrower1Name")

'Updates the Rescind Date
Set companyHolidays1 = Range("CompanyHolidays")
closingDate = Range("SigningDate").Value2
proposedRescindDate = closingDate + 3
netWorkDayResult = (Application.WorksheetFunction.NetworkDays_Intl(closingDate + 1, proposedRescindDate, 11, companyHolidays1))
daysToAdd = 3 - netWorkDayResult
proposedRescindDate = proposedRescindDate + daysToAdd
sundayCheck = (Application.WorksheetFunction.Weekday(proposedRescindDate, 2))
If sundayCheck = 7 Then
    proposedRescindDate = proposedRescindDate + 1
End If
closingDate = proposedRescindDate
Range("RescindDate").Value2 = closingDate

Call ageCalculation 'to update Over65 on all borrowers then this function is called.

'moved to its own method
'If Range("Borrower1FKA") <> 0 Then 'Adds the borrowers FKA to their name if needed.
'    Range("Borrower1Name") = Range("Borrower1Name") & " FKA " & Range("Borrower1FKA")
'End If
'If Range("Borrower2FKA") <> 0 Then
'    Range("Borrower2Name") = Range("Borrower2Name") & " FKA " & Range("Borrower2FKA")
'End If
'If Range("Borrower3FKA") <> 0 Then
'    Range("Borrower3Name") = Range("Borrower3Name") & " FKA " & Range("Borrower3FKA")
'End If

End Sub

Sub calculateTotalAmountDue() 'Calculates the total tax amount due per property.

Dim i As Integer    ' For first for loop
Dim j As Integer    ' For second for loop
Dim nextTotal As String
Dim nextProp As String
Dim totalProp As Single
Dim totalTax As Single

'Updates the Prop(1-25)TotalAmountDue fields
For i = 1 To Range("NumberofProperties")
    nextTotal = "Prop" & i & "TotalAmountDue"
    'Cycles through 1-4 potential amounts in the PropiAmountDuej, adds them to totalProp
    For j = 1 To 4
        nextProp = "Prop" & i & "AmountDue" & j
        totalProp = totalProp + Range(nextProp).Value
    Next j
    Range(nextTotal) = totalProp
    'Updates the totalTax variable to keep a running tab for the TotalTaxAmounts field
    totalTax = totalTax + totalProp
    'resets totalProp so next property doesn't include it.
    totalProp = 0
Next i
Range("TotalTaxAmount") = totalTax

End Sub

Sub findOCCCNumber() 'Index and Match to find the OCCC Number when the Entity is chosen.

Dim WS10 As Worksheet
Dim test1 As Variant
Dim NumberArray1 As Variant
Dim NumberArray2 As Variant

Set WS10 = Worksheets("DropdownInfo")
Set NumberArray1 = WS10.Range("C3:C15")
Set NumberArray2 = WS10.Range("B3:B15")

test1 = Application.Index(NumberArray1, Application.Match(Range("Entity").Value, NumberArray2, 0), 1)

Range("OCCCNumber").Value = test1

End Sub
Sub ageCalculation() 'IF(((TODAY()-Borrower1DOB + 1) / 365.25)<65,"Young - under 65","Old - Over 65")

Dim today As Date
Dim age As Date
Dim number As Double

today = Range("todayDate").Value2

If Range("Borrower1DOB").Value2 = "N/A" Then
    Exit Sub
End If

If Range("Borrower1DOB").Value2 <> 0 Then 'Check borrower1's age
    age = Range("Borrower1DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B1Over65") = "N"
    Else
        Range("B1Over65") = "Y"
    End If
Else
    Range("B1Over65") = "N"
End If

If Range("Borrower2DOB").Value2 <> 0 Then 'Check borrower2's age
    age = Range("Borrower2DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B2Over65") = "N"
    Else
        Range("B2Over65") = "Y"
    End If
Else
    Range("B2Over65") = "N"
End If

If Range("Borrower3DOB").Value2 <> 0 Then 'Check borrower3's age
    age = Range("Borrower3DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B3Over65") = "N"
    Else
        Range("B3Over65") = "Y"
    End If
Else
    Range("B3Over65") = "N"
End If

End Sub
Sub calculateBorrower1Age()

Dim today As Date
Dim age As Date
Dim number As Double

today = Range("todayDate").Value2

If Range("Borrower1DOB").Value2 = "N/A" Then
    Exit Sub
End If

If Range("Borrower1DOB").Value2 <> 0 Then
    age = Range("Borrower1DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B1Over65") = "N"
    Else
        Range("B1Over65") = "Y"
    End If
Else
    Range("B1Over65") = "N"
End If

End Sub

Sub calculateBorrower2Age()

Dim today As Date
Dim age As Date
Dim number As Double

If Range("Borrower2DOB").Value2 = "N/A" Then
    Exit Sub
End If

today = Range("todayDate").Value2
If Range("Borrower2DOB").Value2 <> 0 Then
    age = Range("Borrower2DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B2Over65") = "N"
    Else
        Range("B2Over65") = "Y"
    End If
Else
    Range("B2Over65") = "N"
End If

End Sub

Sub calculateBorrower3Age()

Dim today As Date
Dim age As Date
Dim number As Double

If Range("Borrower3DOB").Value2 = "N/A" Then
    Exit Sub
End If

today = Range("todayDate").Value2

If Range("Borrower3DOB").Value2 <> 0 Then
    age = Range("Borrower3DOB").Value2
    number = (today - age + 1) / 365.25
    If number < 65 Then
        Range("B3Over65") = "N"
    Else
        Range("B3Over65") = "Y"
    End If
Else
    Range("B3Over65") = "N"
End If

End Sub

Sub findNMLS() 'Index and Match to find the NMLS ID when the LoanOfficer is chosen.

Dim WS11 As Worksheet
Dim test1 As Variant
Dim NumberArray1 As Variant
Dim NumberArray2 As Variant

Set WS11 = Worksheets("DropdownInfo")
Set NumberArray1 = WS11.Range("R22:R29")
Set NumberArray2 = WS11.Range("Q22:Q29")

test1 = Application.Index(NumberArray1, Application.Match(Range("LoanOfficer").Value, NumberArray2, 0), 1)
Range("NMLS").Value = test1

End Sub
Sub ClearBorrowers()

Dim FieldsToClear As Variant
Dim x As Variant
Dim Rg1 As Range

If Range("NumberOfBorrowers").Value = 3 Then
    Exit Sub
ElseIf Range("NumberOfBorrowers").Value = 1 Then
    FieldsToClear = Array("B2Info", "B3Info")
ElseIf Range("NumberOfBorrowers").Value = 2 Then
    FieldsToClear = Array("B3Info")
End If

For Each x In FieldsToClear
   For Each Rg1 In Range(x)
    Rg1.MergeArea.ClearContents
   Next
Next x

End Sub
Sub CheckClear() 'Msg box confirming user wants to clear the form.

If MsgBox("Do you want to clear everything?", vbYesNo) = vbNo Then
    Exit Sub
Else
    Call ClearAll
End If

End Sub

Sub ClearAll()

Dim FieldsToClear As Variant
Dim x As Variant
    
Dim Rg1 As Range
    FieldsToClear = Array("B1Info", "B2Info", "B3Info", "LoanInfo", "LendingInfo", "TLTAInfo", _
        "Prop1Info", "Prop2Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", _
        "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", _
        "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info", "MiscInfo", "AmountTotalsInfo", "Notes")

Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.EnableEvents = False

For Each x In FieldsToClear
   For Each Rg1 In Range(x)
    Rg1.MergeArea.ClearContents
   Next
Next x

Range("NumberofBorrowers") = 1
Range("NumberofProperties") = 1
Range("CheckBox01") = False
Range("CheckBox02") = False
Range("CheckBox03") = False
Range("CheckBox04") = False

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True

Calculate
Call UpdateAfterEvents

End Sub
Sub UpdateAfterEvents() ' Needed since events are disabled when loading/clearing all to save on processing time.

Range("NumberOfBorrowers") = Range("NumberofBorrowers")
Range("NumberOfProperties") = Range("NumberofProperties")
Range("LoanType") = Range("LoanType")

End Sub
Sub autofill() 'Auto fills information on the loan info and Tax Lien Transfer App

Dim closingCosts As Double

closingCosts = 900 + (100 * (Range("NumberofProperties").Value2 - 1))

Range("UnderwritingFee") = 50
Range("FloodFee") = 45
Range("AssessmentOfValue") = 130
Range("BankruptcySearch") = 45
Range("InternalTitleSearch") = 50
Range("InternalTitleReview") = 0
Range("CreditAmount") = 0
Range("NotaryFee") = 100
Range("MailingFee") = 19
Range("DocumentPreparationFee") = 0
Range("TitleCurativeFee") = 0
Range("ExternalTitleSearch") = 0
Range("ExternalTitleReview") = 25 * Range("NumberofProperties").Value2
Range("ClosingCosts") = closingCosts
Range("CourtCost") = 0
Range("AbstractFees") = 0
Range("OtherFees") = 0

Range("Homestead") = "N"
Range("AgeDeferral") = "N"
Range("AppliedDisability") = "N"
Range("DisabilityDeferral") = "N"
Range("PendingLawsuits") = "N"
Range("PendingCauseNumber") = "N/A"
Range("MortgageLoan") = "N"
Range("MortgageCompany") = "N/A"
Range("YearFinanced") = "N/A"
Range("Bankruptcy") = "N"
Range("BankruptcyCounty") = "N/A"
Range("ExpectedIncome") = "N"
Range("MakePayments") = "Y"
Range("IntentForPayments") = "Income"

Call calcProcessingFee

End Sub
Sub prefillBorrowerInfo() 'Prefills borrower infomation

Dim listOfBorrowerNames As Variant
Dim i As Integer
Dim j As Integer
Dim currentBorrower As String
Dim combinedBorrower As String

listOfBorrowerNames = Array("Bankrupt", "Over65", "Disabled")

Application.Calculation = xlCalculationManual

For i = 1 To Range("NumberofBorrowers")
    currentBorrower = "B" & i
    For j = 0 To UBound(listOfBorrowerNames)
        combinedBorrower = currentBorrower & listOfBorrowerNames(j)
        Range(combinedBorrower) = "N"
    Next j
Next i
Application.Calculation = xlCalculationAutomatic

End Sub
Sub prefillPropAns() 'prefill property infomation

Dim listOfPropNames As Variant
Dim i As Integer
Dim currentProp As String
Dim combinedProp As String

listOfPropNames = Array("MobileHome", "Attached", "TDHCA", "MortgageHolder", "MortgageWith", "TaxLiens", "TaxLoan", "TaxLoanWith", "LoanDate", "LoanBalance", _
                        "Foreclosure", "FDate", "Lawsuit", "Cause")

Application.Calculation = xlCalculationManual

For i = 1 To Range("NumberofProperties")
    currentProp = "Prop" & i
    
    combinedProp = currentProp & listOfPropNames(0)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(1)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(2)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(3)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(4)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(5)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(6)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(7)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(8)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(9)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(10)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(11)
    Range(combinedProp) = ""
    combinedProp = currentProp & listOfPropNames(12)
    Range(combinedProp) = "N"
    combinedProp = currentProp & listOfPropNames(13)
    Range(combinedProp) = ""
Next i
Application.Calculation = xlCalculationAutomatic

End Sub

Sub calcProcessingFee()

Dim processingFee As Double

processingFee = Range("ClosingCosts") - Range("UnderwritingFee") - Range("FloodFee") - Range("AssessmentOfValue") - Range("BankruptcySearch") - _
                Range("InternalTitleSearch") - Range("InternalTitleReview") - Range("CreditAmount") - Range("NotaryFee") - Range("MailingFee") - _
                Range("DocumentPreparationFee") - Range("TitleCurativeFee") - Range("ExternalTitleSearch") - Range("ExternalTitleReview")
Range("ProcessingFee") = processingFee

Call UpdateTotals

End Sub


Sub ClearProperties()
Dim FieldsToClear As Variant
Dim x As Variant
Dim Rg1 As Range

Application.Calculation = xlManual
Application.ScreenUpdating = False

If Range("NumberofProperties").Value = 25 Then
    Exit Sub
    
ElseIf Range("NumberOfProperties").Value = 1 Then
    FieldsToClear = Array("Prop2Info", "Prop3Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", _
    "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", _
    "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 2 Then
    FieldsToClear = Array("Prop3Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", "Prop10Info", _
    "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", _
    "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 3 Then
    FieldsToClear = Array("Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", "Prop10Info", "Prop11Info", _
    "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", _
    "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 4 Then
    FieldsToClear = Array("Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", "Prop10Info", "Prop11Info", "Prop12Info", _
    "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", _
    "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 5 Then
    FieldsToClear = Array("Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", _
    "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", _
    "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 6 Then
    FieldsToClear = Array("Prop7Info", "Prop8Info", "Prop9Info", "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", _
    "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 7 Then
    FieldsToClear = Array("Prop8Info", "Prop9Info", "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", _
    "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 8 Then
    FieldsToClear = Array("Prop9Info", "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", _
    "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 9 Then
    FieldsToClear = Array("Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", _
    "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 10 Then
    FieldsToClear = Array("Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", _
    "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 11 Then
    FieldsToClear = Array("Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", _
    "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
    
ElseIf Range("NumberOfProperties").Value = 12 Then
    FieldsToClear = Array("Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", _
    "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 13 Then
    FieldsToClear = Array("Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", _
    "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 14 Then
    FieldsToClear = Array("Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", _
    "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 15 Then
    FieldsToClear = Array("Prop16Info", "Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", _
    "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 16 Then
    FieldsToClear = Array("Prop17Info", "Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")

ElseIf Range("NumberOfProperties").Value = 17 Then
    FieldsToClear = Array("Prop18Info", "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 18 Then
    FieldsToClear = Array("Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 19 Then
    FieldsToClear = Array("Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")

ElseIf Range("NumberOfProperties").Value = 20 Then
    FieldsToClear = Array("Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")

ElseIf Range("NumberOfProperties").Value = 21 Then
    FieldsToClear = Array("Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info")

ElseIf Range("NumberOfProperties").Value = 22 Then
    FieldsToClear = Array("Prop23Info", "Prop24Info", "Prop25Info")
    
ElseIf Range("NumberOfProperties").Value = 23 Then
    FieldsToClear = Array("Prop24Info", "Prop25Info")

ElseIf Range("NumberOfProperties").Value = 24 Then
    FieldsToClear = Array("Prop25Info")
End If

For Each x In FieldsToClear
   For Each Rg1 In Range(x)
    Rg1.MergeArea.ClearContents
   Next
Next x

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
