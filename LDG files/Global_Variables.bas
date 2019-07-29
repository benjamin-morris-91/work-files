Attribute VB_Name = "Global_Variables"
Option Explicit

'List all global variables below
Global loanDocFile As String
Global databaseFile As String
Global mergeFieldsFile As String
Global individualMailMergeFile As String
Global corporateMailMergeFile As String
Global saveLocation As String
Global individualMailMergeCert As String
Global GFEMergeForm As String
Global LoanDocWB As Workbook
Global LoanDocSht1 As Worksheet
Global LoanDocDB As Worksheet
Global LoanDocNDB As Worksheet


Sub assignFileNames() 'Call UserFilePath 'Updates the correct location of the loan docs file

loanDocFile = Range("PathToLoanDoc")
databaseFile = Range("PathToDatabase")
mergeFieldsFile = Range("PathToMergeFields")
individualMailMergeFile = Range("PathToIndividualMailMerge")
GFEMergeForm = Range("PathToGFE")
corporateMailMergeFile = Range("PathToCorporateMailMerge")
saveLocation = Range("PathToSaveLocation")
individualMailMergeCert = Range("PathToCert")

Set LoanDocWB = Workbooks.Open(loanDocFile)
'Set LoanDocSht1 = LoanDocWB.Sheets("Sheet1")
'Set LoanDocDB = LoanDocWB.Sheets("Database")
'Set LoanDocNDB = LoanDocWB.Sheets("NewDatabase")
Set LoanDocSht1 = Worksheets("Sheet1")
Set LoanDocDB = Worksheets("Database")
Set LoanDocNDB = Worksheets("NewDatabase")

End Sub
