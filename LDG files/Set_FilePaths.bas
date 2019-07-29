Attribute VB_Name = "Set_FilePaths"
Option Explicit

Sub UserFilePath()

ActiveWorkbook.Worksheets("DropdownInfo").Range("PathToLoanDoc").Value = ActiveWorkbook.FullName
ActiveWorkbook.Worksheets("DropdownInfo").Range("PathToSaveLocation").Value = CreateObject("wscript.shell").SpecialFolders(4) & "\" 'Finds desktop location for the current user.

End Sub


Sub UpdateSaveLocation() 'Allows users to update the save location for the loan docs.

Dim str As String
Dim newSaveLocation As Variant

str = CreateObject("wscript.shell").SpecialFolders(4) & "\" 'Finds desktop location for the current user.
newSaveLocation = InputBox("Enter where you would like the files saved. Make sure to end path with a  ' \'", "Update Save Location", str)

If newSaveLocation = "" Then 'Exits sub if user clicks cancel or enters a blank answer.
    Exit Sub
End If

Range("PathToSaveLocation") = newSaveLocation

saveLocation = Range("PathToSaveLocation")


End Sub

Sub UpdateFileLocation()

Dim str As String

str = InputBox("Please enter a new location where the files are stored. Make sure to end path with a  ' \' ")

If str = "" Then 'Exits sub if user clicks cancel or enters a blank answer.
    Exit Sub
End If

Range("PathToIndividualMailMerge") = str & "IndividualDocumentsMergeFile.docx"
Range("PathToCorporateMailMerge") = str & "CorporateDocumentsMergeFile.docx"
Range("PathToCert") = str & "IndividualCertifiedStatementTTT.docx"
Range("PathToGFE") = str & "DisclosuresMergeForm.docx"

End Sub

Sub UpdateDatabaseMergeLocation()

Dim str As String

str = InputBox("Please enter a new location where the database and merge fields are stored. Make sure to end path with a  ' \' ")

If str = "" Then 'Exits sub if user clicks cancel or enters a blank answer.
    Exit Sub
End If

Range("PathToDatabase") = str & "Loan Database.xlsx"
Range("PathToMergeFields") = str & "MergeFields.csv"

End Sub

Sub PointToWDrive()

Dim str As String
str = "\\zeus\hkdata\LDG Files\Files\"

Range("PathToIndividualMailMerge") = str & "IndividualDocumentsMergeFile.docx"
Range("PathToCorporateMailMerge") = str & "CorporateDocumentsMergeFile.docx"
Range("PathToCert") = str & "IndividualCertifiedStatementTTT.docx"
Range("PathToGFE") = str & "DisclosuresMergeForm.docx"
Range("PathToDatabase") = str & "Loan Database.xlsx"
Range("PathToMergeFields") = str & "MergeFields.csv"

Worksheets("Database").Range("G24").Value = "Zeus Drive"

End Sub

Sub PointToBensDesktop()

Dim str As String
str = "C:\Users\Ben\Desktop\Testing2\"

Range("PathToIndividualMailMerge") = str & "IndividualDocumentsMergeFile.docx"
Range("PathToCorporateMailMerge") = str & "CorporateDocumentsMergeFile.docx"
Range("PathToCert") = str & "IndividualCertifiedStatementTTT.docx"
Range("PathToGFE") = str & "DisclosuresMergeForm.docx"
Range("PathToDatabase") = str & "Loan Database.xlsx"
Range("PathToMergeFields") = str & "MergeFields.csv"

Worksheets("Database").Range("G24").Value = "Desktop Folder"

End Sub
