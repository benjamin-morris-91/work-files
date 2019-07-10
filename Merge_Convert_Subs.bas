Attribute VB_Name = "Merge_Convert_Subs"
Option Explicit
Dim flag1 As Integer 'Used to determine if files get overwritten or not. 1 is true (overwrite) and 0 is false (end sub)
Dim inputSelection As Variant

Sub DoMailMergeDocs()
'************* Creates new word document based off of what is in the MergeField document. Names it LOANNUMBER + BORROWER1NAME + .docx ********************

Dim WordApp As Object
Dim WordDoc As Object        'Master Individual Documents
Dim strWorkbookName As String
Dim FileName As String
Dim pathtosave As String

FileName = Range("LoanNumber") & " " & Range("Borrower1Name") & ".docx" 'Sets the name of the file as the loan number + name
flag1 = 1
pathtosave = saveLocation & FileName

    If Dir(pathtosave, 0) <> vbNullString Then 'Check to see if the file exists and if the user wants to overwrite
        inputSelection = MsgBox("File already exists. Would you like to replace?", vbQuestion + vbYesNo)
        
        If inputSelection = vbYes Then 'Overwrites the file
            flag1 = 1
        Else
            flag1 = 0
        End If
    End If
    
    If flag1 = 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
    If WordApp Is Nothing Then  'Set WordApp = GetObject(, "Word.Application") 'WordApp
        Set WordApp = CreateObject("Word.Application") 'WordApp
    Else
        MsgBox "Exit all Word windows and try again."
        Exit Sub
    End If
    On Error GoTo 0
    
    Set WordDoc = WordApp.Documents.Open(individualMailMergeFile) 'Path of Master Word doc here

    WordDoc.MailMerge.MainDocumentType = wdFormLetters
    WordDoc.MailMerge.OpenDataSource _
            Name:=mergeFieldsFile, _
            AddToRecentFiles:=False, _
            Revert:=False, _
            Format:=wdOpenFormatAuto, _
            Connection:="Data Source=" & mergeFieldsFile & ";Mode=Read", _
            SQLStatement:="SELECT * FROM 'Sheet1$'"

    With WordDoc.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        .Execute Pause:=False
    End With
 
    WordApp.ActiveDocument.SaveAs2 pathtosave, wdFormatDocumentDefault   'saves the file
 
    WordApp.Visible = False
    WordApp.ActiveDocument.Close SaveChanges:=False
    
    WordDoc.Close SaveChanges:=False
    WordApp.Quit
    
    Set WordDoc = Nothing
    Set WordApp = Nothing
    
    Call AdjustTables

End Sub

Sub DoMailGFE()
'************* Creates new word document based off of what is in the MergeField document. Names it LOANNUMBER + BORROWER1NAME + .docx ********************
'Creates word doc from mail merge and saves a copy as a pdf.

If flag1 = 0 Then 'Check the flag to see if the sub should exit.
    Exit Sub
End If

Dim wd As Object
Dim wdocSource As Object        'Master Individual Documents
Dim strWorkbookName As String
Dim fileNameForGFE As String
Dim pathtosave As String
Dim fileNameForCert As String
Dim filePath As String
Dim updatedFileName As String

fileNameForGFE = Range("LoanNumber") & " Disclosures.docx" 'Sets the name of the file as the loan number + name

    On Error Resume Next
        If wd Is Nothing Then
            Set wd = CreateObject("Word.Application")
        Else
            MsgBox "Exit all Word windows and try again."
            Exit Sub
        End If
    On Error GoTo 0
    
    Set wdocSource = wd.Documents.Open(GFEMergeForm) 'Path of Master Word doc here

    strWorkbookName = mergeFieldsFile     'Name and Path of MergeFields doc here

    wdocSource.MailMerge.MainDocumentType = wdFormLetters

    wdocSource.MailMerge.OpenDataSource _
            Name:=strWorkbookName, _
            AddToRecentFiles:=False, _
            Revert:=False, _
            Format:=wdOpenFormatAuto, _
            Connection:="Data Source=" & strWorkbookName & ";Mode=Read", _
            SQLStatement:="SELECT * FROM 'Sheet1$'"

    With wdocSource.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        .Execute Pause:=False
    End With

    pathtosave = saveLocation & fileNameForGFE

    wd.ActiveDocument.SaveAs2 pathtosave, wdFormatDocumentDefault
    
    wd.Visible = False
    wdocSource.Close SaveChanges:=False

    Set wdocSource = Nothing

fileNameForCert = Range("LoanNumber") & " Disclosures.docx"
filePath = saveLocation & fileNameForCert
updatedFileName = Range("LoanNumber") & " Disclosures"

    wd.Documents.Open filePath 'File path that will be opened and converted to PDF
    wd.Visible = False
    wd.ActiveDocument.ExportAsFixedFormat OutputFileName:=saveLocation & updatedFileName & ".pdf", ExportFormat:=wdExportFormatPDF 'Converts the active document to a pdf
    wd.Quit
    Set wd = Nothing

End Sub


Sub DoMailMergeCert()
'************* Creates new word document based off of what is in the MergeField document. Names it LOANNUMBER + BORROWER1NAME + .docx ********************
'Creates mail merge form and converts it to pdf

'Check the flag to see if the sub should exit.
If flag1 = 0 Then
    Exit Sub
End If

Dim wd As Object
Dim wdocSource As Object        'Master Individual Documents
Dim strWorkbookName As String
Dim fileNameForCert As String
Dim pathtosave As String
Dim filePath As String
Dim updatedFileName As String

fileNameForCert = Range("LoanNumber") & " Cert.docx" 'Sets the name of the file as the loan number + name

    On Error Resume Next
        If wd Is Nothing Then
            Set wd = CreateObject("Word.Application")
        Else
            MsgBox "Exit all Word windows and try again."
            Exit Sub
        End If
    On Error GoTo 0
    
    Set wdocSource = wd.Documents.Open(individualMailMergeCert)    'Path of Master Word doc here

    strWorkbookName = mergeFieldsFile 'Name and Path of MergeFields doc here

    wdocSource.MailMerge.MainDocumentType = wdFormLetters

    wdocSource.MailMerge.OpenDataSource _
            Name:=strWorkbookName, _
            AddToRecentFiles:=False, _
            Revert:=False, _
            Format:=wdOpenFormatAuto, _
            Connection:="Data Source=" & strWorkbookName & ";Mode=Read", _
            SQLStatement:="SELECT * FROM 'Sheet1$'"

    With wdocSource.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        .Execute Pause:=False
    End With

    pathtosave = saveLocation & fileNameForCert

    wd.ActiveDocument.SaveAs2 pathtosave, wdFormatDocumentDefault
    
    wd.Visible = False
    wdocSource.Close SaveChanges:=False

    Set wdocSource = Nothing

filePath = saveLocation & fileNameForCert
updatedFileName = Range("LoanNumber") & " Cert"
wd.Documents.Open filePath 'File path that will be opened and converted to PDF
wd.Visible = False

wd.ActiveDocument.ExportAsFixedFormat OutputFileName:=saveLocation & updatedFileName & ".pdf", ExportFormat:=wdExportFormatPDF 'Converts the active document to a pdf

wd.Quit
Set wd = Nothing

End Sub

Sub ConvertToPDF()

If flag1 = 0 Then 'Check the flag to see if the sub should exit.
    Exit Sub
End If

Dim WDpdf As Object
Dim FileName As String
Dim filePath As String
Dim updatedFileName As String

On Error Resume Next
    If WDpdf Is Nothing Then
        Set WDpdf = CreateObject("Word.Application")
    End If
    On Error GoTo 0

FileName = Range("LoanNumber") & " " & Range("Borrower1Name") & ".docx"
filePath = saveLocation & FileName
updatedFileName = Range("LoanNumber") & " " & Range("Borrower1Name")
 
WDpdf.Visible = False
        
WDpdf.Visible = False
WDpdf.Documents.Open filePath

WDpdf.ActiveDocument.ExportAsFixedFormat OutputFileName:=saveLocation & updatedFileName & ".pdf", ExportFormat:=wdExportFormatPDF 'Converts the active document to a pdf

WDpdf.Quit
Set WDpdf = Nothing

End Sub

Sub StandAloneGFE() 'Runs when user wants to generate JUST the GFE docs.

Dim wd As Object
Dim wdocSource As Object        'Master Individual Documents
Dim strWorkbookName As String
Dim fileNameForGFE As String
Dim pathtosave As String
Dim fileNameForCert As String
Dim filePath As String
Dim updatedFileName As String

fileNameForGFE = Range("LoanNumber") & " GFE.docx" 'Sets the name of the file as the loan number + name

    On Error Resume Next
        If wd Is Nothing Then
            Set wd = CreateObject("Word.Application")
        Else
            MsgBox "Exit all Word windows and try again."
            Exit Sub
        End If
    On Error GoTo 0
    
    Set wdocSource = wd.Documents.Open(GFEMergeForm) 'Path of Master Word doc here

    strWorkbookName = mergeFieldsFile     'Name and Path of MergeFields doc here

    wdocSource.MailMerge.MainDocumentType = wdFormLetters

    wdocSource.MailMerge.OpenDataSource _
            Name:=strWorkbookName, _
            AddToRecentFiles:=False, _
            Revert:=False, _
            Format:=wdOpenFormatAuto, _
            Connection:="Data Source=" & strWorkbookName & ";Mode=Read", _
            SQLStatement:="SELECT * FROM 'Sheet1$'"

    With wdocSource.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = wdDefaultFirstRecord
            .LastRecord = wdDefaultLastRecord
        End With
        .Execute Pause:=False
    End With

    pathtosave = saveLocation & fileNameForGFE

    wd.ActiveDocument.SaveAs2 pathtosave, wdFormatDocumentDefault
    
    wd.Visible = False
    wdocSource.Close SaveChanges:=False

    Set wdocSource = Nothing

fileNameForCert = Range("LoanNumber") & " GFE.docx"
filePath = saveLocation & fileNameForCert
updatedFileName = Range("LoanNumber") & " GFE"

    wd.Documents.Open filePath 'File path that will be opened and converted to PDF
    wd.Visible = False
    wd.ActiveDocument.ExportAsFixedFormat OutputFileName:=saveLocation & updatedFileName & ".pdf", ExportFormat:=wdExportFormatPDF 'Converts the active document to a pdf
    wd.Quit
    Set wd = Nothing
    
End Sub

Sub AdjustTables()

Dim WordApp As Object
Dim WordDoc As Object
Dim newRow As ListRow
Dim tb1 As ListObject
Dim FileName As String
Dim filePath As String
Dim i As Integer
Dim ii As Integer

FileName = Range("LoanNumber") & " " & Range("Borrower1Name") & ".docx"
filePath = saveLocation & FileName

    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    Set WordDoc = WordApp.Documents.Open(filePath)
 
    If Range("NumberOfUniqueEntities") > 2 Then 'Change "number" to "NumberOfProperties"
        ii = 5  'Table we want to add is the fifth table in the word doc.
        For i = 3 To Range("NumberOfUniqueEntities")
            'ii = 11 + Range("NumberOfSworns")   '11 is base case, add that to the number of sworns generated.
            WordDoc.Tables.Item(ii).Rows.Add (WordDoc.Tables.Item(ii).Rows.Item(i))                 'Add the row
            WordDoc.Tables.Item(ii).Cell(i, 1).Range.Text = "to " & Range("UniqueEntity" & i)       'Fill it with data
            WordDoc.Tables.Item(ii).Cell(i, 2).Range.Text = "$" & Range("UniqueAmount" & i)         'CONVERT STRING TO DECIMAL
        Next
    End If
    
WordDoc.Close SaveChanges:=True
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing

End Sub
