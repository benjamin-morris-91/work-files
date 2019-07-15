Attribute VB_Name = "Save_Load_Subs"
Option Explicit

Sub LoadFromDatabase() 'runs when user clicks the "Load Loan" button and also during the save phase.

'**********************************************
'Add some error handeling to cover if the match returns a N/A
'**********************************************

    Dim Wb1 As Workbook                     'Loan Doc
    Dim Wb2 As Workbook                     'Database Doc
    Dim newDatabase As Worksheet            'copying from main file
    Dim copyDatabaseFile As Worksheet       'pasting into database file
    
    Call assignFileNames
    
    Set Wb1 = Workbooks.Open(loanDocFile) 'This is the Loan Doc Generator file
    Set newDatabase = Worksheets("NewDatabase")
    Set Wb2 = Workbooks.Open(databaseFile) 'This is the database file
    Set copyDatabaseFile = Worksheets("Sheet1")
    
    newDatabase.Cells.ClearContents
    copyDatabaseFile.Range("E1") = LoanDocDB.Range("LNPlusName") 'Copy from the LoanUI to the database file range E1.
    copyDatabaseFile.Range("F1") = Application.Match(copyDatabaseFile.Range("E1"), copyDatabaseFile.Range("B:B"), 0) 'Provides row to copy/replace
    
    copyDatabaseFile.Cells.Copy
    newDatabase.Range("A1").PasteSpecial xlPasteValues

    Application.CutCopyMode = False
    Wb2.Close SaveChanges:=True
    
    Set copyDatabaseFile = Nothing
    Set newDatabase = Nothing
    
    Set Wb1 = Nothing
    
End Sub

Sub LoadToDatabase() 'Happens during the save phase

    Dim Wb1 As Workbook
    Dim Wb4 As Workbook
    Dim newDatabase As Worksheet
    Dim copyDatabaseFile As Worksheet
    
    Set Wb1 = Workbooks.Open(loanDocFile) 'This is the LoanUI file
    Set newDatabase = Worksheets("NewDatabase")
    Set Wb4 = Workbooks.Open(databaseFile) 'This is the database file
    Set copyDatabaseFile = Worksheets("Sheet1")
    copyDatabaseFile.Range("F1") = newDatabase.Range("F1")

    copyDatabaseFile.Rows(copyDatabaseFile.Range("F1")).ClearContents
    newDatabase.Rows(newDatabase.Range("F1")).Copy
    copyDatabaseFile.Rows(copyDatabaseFile.Range("F1")).PasteSpecial xlPasteValues
    copyDatabaseFile.Range("A1").Select
    
    Wb4.Close SaveChanges:=True
    Set copyDatabaseFile = Nothing
    
    Application.CutCopyMode = False
    Set newDatabase = Nothing
    
    Set Wb1 = Nothing
    
End Sub

Sub SaveLoan()

    
    'Dim inputSelection As Variant
    Dim rowNumberToReplace As Integer
    
    Call CheckBeforeGenerating
    If Range("LoanReady") = "No" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Call assignFileNames
    Application.ScreenUpdating = False
    Call UpdateTotals
    Application.ScreenUpdating = False
    Call UniqueLoans
    Application.ScreenUpdating = False
    Call SolvePayment
    Application.ScreenUpdating = False
    Call ClearBorrowers
    Application.ScreenUpdating = False
    Call ClearProperties
    Application.ScreenUpdating = False
    Call LoadFromDatabase
    Application.ScreenUpdating = False
    Call AddDupes
    Application.ScreenUpdating = False
    Call splitLegals
    Application.ScreenUpdating = False
    
    
    'Searches current loan numbers in the NewDatabase sheet for duplicates, enters if statement if found.
    If Not IsError(Application.Match(LoanDocDB.Range("C6"), LoanDocNDB.Range("A:A"), 0)) Then
        
        SaveUF.Show

        If Range("SaveChoice") = 1 Then 'Will overwrite the data if the user selects 1
            rowNumberToReplace = LoanDocNDB.Range("F1").Value
            LoanDocDB.Range("AppInfo").Copy
            LoanDocNDB.Rows(rowNumberToReplace).PasteSpecial
            
            Call LoadToDatabase
        End If
    Else 'No Duplicates so copies Range("AppInfo") to a new line in the NewDatabase sheet.
        LoanDocDB.Range("AppInfo").Copy
        LoanDocNDB.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        LoanDocNDB.Range("F1") = Application.Match(LoanDocDB.Range("C6"), LoanDocNDB.Range("A:A"), 0)
        
        Call LoadToDatabase
        MsgBox "New Record saved."
    End If
    
    Worksheets("Sheet1").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call SplitSSN
    Call TrackChanges
    
End Sub

Sub LoadLoan() 'Loads a loan into the worksheet based off a user's input of the loan number

    Call assignFileNames
    
    Worksheets("NewDatabase").Activate
    Worksheets("NewDatabase").Range("E1") = Range("LNToLoad")
    
    Worksheets("NewDatabase").Range("D1:F1").Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    If Range("TypeSelected") = "Top" Then
        Worksheets("NewDatabase").Range("F1") = Application.Match(Worksheets("NewDatabase").Range("E1"), Worksheets("NewDatabase").Range("B:B"), 0)
    Else
        Worksheets("NewDatabase").Range("F1") = Application.Match(Worksheets("NewDatabase").Range("E1"), Worksheets("NewDatabase").Range("BL:BL"), 0)
    End If

    If IsError(Worksheets("NewDatabase").Range("F1")) Then
        MsgBox "No Match found. Please try again."
    ElseIf Range("TypeSelected") = "Top" Then 'Checks the TypeSelected range to see what kind of match should run.
        Rows(Application.Match(Worksheets("NewDatabase").Range("E1"), Worksheets("NewDatabase").Range("B:B"), 0)).Copy
        Worksheets("Database").Activate
        Worksheets("Database").Range("A10").PasteSpecial xlPasteValues
    Else
        Rows(Application.Match(Worksheets("NewDatabase").Range("E1"), Worksheets("NewDatabase").Range("BC:BC"), 0)).Copy
        Worksheets("Database").Activate
        Worksheets("Database").Range("A10").PasteSpecial xlPasteValues
    End If
    
    Call ClearAll
    Call Replace
    Call SplitSSN
    
    Worksheets("Sheet1").Activate

End Sub
Sub Replace()

    Dim Array1 As Variant
    Dim Array2 As Variant
    Dim rng1 As Range
    Dim rng2 As Range
    Dim i As Long

    Worksheets("Database").Activate
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
   
    Set rng1 = Worksheets("Database").Range("A1", Range("A1").End(xlToRight)) 'Selects a row (with more than one element) no matter the size.
    Set rng2 = Worksheets("Database").Range("A10", Range("A10").End(xlToRight))

    Array1 = rng1.Value2
    Array2 = rng2.Value2

    For i = 1 To Range("LengthOfArray").Value
        Range(Array1(1, i)) = Array2(1, i)
    Next
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Calculate
    Call UpdateAfterEvents
    
End Sub
Sub GenerateForms()

    Dim Wb12 As Workbook
    Dim mergeFieldsWS As Worksheet
    Dim b1Name As String
    Dim b2Name As String
    Dim b3Name As String
    
    Call CheckBeforeGenerating
    If Range("LoanReady") = "No" Then
        Exit Sub
    End If
    
    b1Name = Range("Borrower1Name")
    b2Name = Range("Borrower2Name")
    b3Name = Range("Borrower3Name")

    Application.ScreenUpdating = False
    
    Call SaveLoan
    Call FKACheck
    'Call UpdateFlags
    Call calcProcessingFee
    'Call LegalLoop 'Checks the length of the legal description before generating docs.
    If Range("LegalCriteria") = False Then
        Exit Sub
    End If
 
    If FSO.FileExists(mergeFieldsFile) = True Then 'Checks to see if the CSV file is present before creating a new version
        FSO.DeleteFile (mergeFieldsFile)
    End If
    
    Worksheets("Database").Activate
    LoanDocDB.Range("A1", Range("A2").End(xlToRight)).Copy
    
    Set Wb12 = Workbooks.Add
    Set mergeFieldsWS = Wb12.Worksheets("Sheet1")

    mergeFieldsWS.Range("A1").PasteSpecial xlPasteValues
    mergeFieldsWS.Range("A1").PasteSpecial xlPasteFormats
    
    Wb12.SaveAs FileName:=mergeFieldsFile, FileFormat:=xlCSVUTF8, CreateBackup:=False

    Application.CutCopyMode = False
    Wb12.Close SaveChanges:=True
    
    Set mergeFieldsWS = Nothing
    Set Wb12 = Nothing
    Worksheets("Sheet1").Activate

'*****************************
'    Application.StatusBar = "Working... 0% Completed"
'    Call DoMailMergeDocs 'Mail merge of individual docs
'    Application.StatusBar = "Working... 50% Completed"
'    Call ConvertToPDF 'Converts individual docs to pdf
'    Application.StatusBar = "Working... 80% Completed"
    
'Took out 7/8/19 against my better judgement. BM
'    Call DoMailGFE ' Creates word doc and pdf
'    Application.StatusBar = "Working... 80% Completed"


'    Call DoMailMergeCert ' Creates word doc and pdf
'    Application.StatusBar = "Finished!"
'*****************************
    Range("Borrower1Name") = b1Name
    Range("Borrower2Name") = b2Name
    Range("Borrower3Name") = b3Name
    Call FKASwap
    
    Application.ScreenUpdating = True
    
    'Delete CSV file here
    'FSO.DeleteFile (mergeFieldsFile)
    'FSO.DeleteFile ("W:\Marketing\Loan Doc Templates\BM Save Folder\MergeFields.cvs")
    
End Sub

Sub GenerateGFE()

    Dim Wb12 As Workbook
    Dim mergeFieldsWS As Worksheet
    Dim b1Name As String
    Dim b2Name As String
    Dim b3Name As String
    
    If Range("LoanReady") = "No" Then
        Exit Sub
    End If
    
    b1Name = Range("Borrower1Name")
    b2Name = Range("Borrower2Name")
    b3Name = Range("Borrower3Name")

    Application.ScreenUpdating = False
    
    Call SaveLoan
    Call FKACheck
    Call UpdateFlags
    Call calcProcessingFee
    'Call LegalLoop 'Checks the length of the legal description before generating docs.
    
    If Range("LegalCriteria") = False Then
        Exit Sub
    End If

    If FSO.FileExists(mergeFieldsFile) = True Then 'Checks to see if the CSV file is present before creating a new version
        FSO.DeleteFile (mergeFieldsFile)
    End If
    
    Worksheets("Database").Activate
    LoanDocDB.Range("A1", Range("A2").End(xlToRight)).Copy

    
    Set Wb12 = Workbooks.Add
    Set mergeFieldsWS = Wb12.Worksheets("Sheet1")

    mergeFieldsWS.Range("A1").PasteSpecial xlPasteValues
    mergeFieldsWS.Range("A1").PasteSpecial xlPasteFormats
    
    Wb12.SaveAs FileName:=mergeFieldsFile, FileFormat:=xlCSVUTF8, CreateBackup:=False

    Application.CutCopyMode = False
    Wb12.Close SaveChanges:=True
    
    Set mergeFieldsWS = Nothing
    Set Wb12 = Nothing
    Worksheets("Sheet1").Activate

'*****************************
    
    Call StandAloneGFE ' Creates word doc and pdf

'*****************************
    Range("Borrower1Name") = b1Name
    Range("Borrower2Name") = b2Name
    Range("Borrower3Name") = b3Name
    Call FKASwap
    
    Application.ScreenUpdating = True
    
    'Delete CSV file here
    'FSO.DeleteFile (mergeFieldsFile)
    'FSO.DeleteFile ("W:\Marketing\Loan Doc Templates\BM Save Folder\MergeFields.cvs")
    
End Sub

Sub CheckBeforeGenerating()

Call UpdateTotals

If Range("LoanNumber") = "" Then
    Range("LoanReady") = "No"
    MsgBox "Need to have a loan number entered before generating or saving docs"
ElseIf Range("Borrower1Name") = "" Then
    Range("LoanReady") = "No"
    MsgBox "Need to have at least one borrower entered before generating or saving docs"
ElseIf Range("Prop1Address") = "" Then
    Range("LoanReady") = "No"
    MsgBox "Need to have at least one property entered before generating or saving docs"
ElseIf Range("AmountToTaxCollector") < 1 Then
    Range("LoanReady") = "No"
    MsgBox "You need to have a loan greater than $0 before generating or saving docs"
Else
    Range("LoanReady") = "Yes"
End If


End Sub

Sub CheckBeforeSaving()

If Range("LoanNumber") = "" Then
    Range("SaveReady") = "No"
    MsgBox "Need to have a loan number entered before saving docs"
ElseIf Range("Borrower1Name") = "" Then
    Range("SaveReady") = "No"
    MsgBox "Need to have at least one borrower entered before saving docs"
ElseIf Range("Prop1Address") = "" Then
    Range("SaveReady") = "No"
    MsgBox "Need to have at least one property entered before saving docs"
Else
    Range("SaveReady") = "Yes"
End If


End Sub

Sub TrackChanges()

Dim Wb2 As Workbook                     'Database Doc
Dim newDatabase As Worksheet            'copying from main file
Dim UpdateDatabaseFile As Worksheet       'pasting into database file
Dim str As String
Dim nextSpot As Range
Dim currentDate As Date
Dim currentMonth As Integer

Call assignFileNames

Set Wb2 = Workbooks.Open(databaseFile) 'This is the database file
Set UpdateDatabaseFile = Worksheets("TrackChanges")

currentDate = Now()
currentMonth = Month(currentDate)

str = FileDateTime(databaseFile) & " - " & Application.UserName
Set nextSpot = UpdateDatabaseFile.Columns(currentMonth).End(xlDown).Offset(1)
nextSpot = str

Wb2.Close SaveChanges:=True
Set UpdateDatabaseFile = Nothing

End Sub

