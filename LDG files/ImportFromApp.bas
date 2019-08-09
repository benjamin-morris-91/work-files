Attribute VB_Name = "ImportFromApp"
Option Explicit

Sub confirmedImport()
    Dim str As String

    Call CheckClear
    str = GetFileName
    Call ImportFromApp

End Sub

Function GetFileName() 'When called, returns the file path that a user selects from the dialog box.
    
    Dim lngCount As Long
    Dim cl As Range

    Set cl = Range("FileName")
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        
        For lngCount = 1 To .SelectedItems.Count
            cl.Worksheet.Hyperlinks.Add _
                Anchor:=cl, Address:=.SelectedItems(lngCount), _
                TextToDisplay:=.SelectedItems(lngCount)
        Next lngCount
        GetFileName = cl 'Returns what the user selected. Esentially the return statement.
    End With
End Function

Sub ExportToApp()

    Dim arrayNameRange As Range
    Dim range2 As Range '
    Dim arrayOne As Variant
    Dim i As Integer
    Dim arrayLength
    Dim wrdApp As Object
    Dim wrdDoc As Object
    Dim templateFile As String
    Dim saveFileName As String
    
    Call assignFileNames
    
'1) Copy necessary data onto row 2 of ApplicationImport tab - Done 7-31-19
'2) Loop through entire row and insert each column into the appropriate bookmark
'3) Save word doc as new file

'1)
'ApplicationImport tab: Put Row 1 into an array, have row 2 be =A1...
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Worksheets("ApplicationImport").Activate

    arrayLength = Range("C4").Value2
    ReDim arrayOne(0 To arrayLength)
    
    Set arrayNameRange = Worksheets("ApplicationImport").Range("A1", Range("A1").End(xlToRight))
    Set range2 = Worksheets("ApplicationImport").Range("A2:EB2")
    range2.ClearContents
    saveFileName = Range("PathToSaveLocation") + "New HK App - " + Range("Borrower1Name")
    
    For i = 0 To (Range("C4") - 1) 'This loop assigns the cells names in row 1 to the array
        arrayOne(i) = arrayNameRange(i + 1)
    Next i

    For i = 0 To (Range("C4") - 1) 'This loop puts the values in the arrayOne array into row 2
        range2(i + 1) = Range(arrayOne(i))
    Next i
    
    
'2)
'Create a new word object like example below
'Point it to the HK loan template
'Emulate the Replace macro in this workbook to reverse the process of putting what is in row 2 back into word

    templateFile = appTemplate 'PathToAppTemplate from global variable
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = False
    Set wrdDoc = wrdApp.Documents.Open(templateFile)
    
    '77 times, for loop doesn't work...
    '1-10
    wrdApp.ActiveDocument.FormFields(arrayOne(0)).Result = range2(1)
    wrdApp.ActiveDocument.FormFields(arrayOne(1)).Result = range2(2)
    wrdApp.ActiveDocument.FormFields(arrayOne(2)).Result = range2(3)
    wrdApp.ActiveDocument.FormFields(arrayOne(3)).Result = range2(4)
    wrdApp.ActiveDocument.FormFields(arrayOne(4)).Result = range2(5)
    wrdApp.ActiveDocument.FormFields(arrayOne(5)).Result = range2(6)
    wrdApp.ActiveDocument.FormFields(arrayOne(6)).Result = range2(7)
    wrdApp.ActiveDocument.FormFields(arrayOne(7)).Result = range2(8)
    wrdApp.ActiveDocument.FormFields(arrayOne(8)).Result = range2(9)
    wrdApp.ActiveDocument.FormFields(arrayOne(9)).Result = range2(10)
'    '11-20
    wrdApp.ActiveDocument.FormFields(arrayOne(10)).Result = range2(11)
    wrdApp.ActiveDocument.FormFields(arrayOne(11)).Result = range2(12)
    wrdApp.ActiveDocument.FormFields(arrayOne(12)).Result = range2(13)
    wrdApp.ActiveDocument.FormFields(arrayOne(13)).Result = range2(14)
    wrdApp.ActiveDocument.FormFields(arrayOne(14)).Result = range2(15)
    If Range("Numberofborrowers") > 1 Then 'fixes a strange error of a wrong date passing through.
        wrdApp.ActiveDocument.FormFields(arrayOne(15)).Result = range2(16)
        wrdApp.ActiveDocument.FormFields(arrayOne(16)).Result = range2(17)
        wrdApp.ActiveDocument.FormFields(arrayOne(17)).Result = range2(18)
        wrdApp.ActiveDocument.FormFields(arrayOne(18)).Result = range2(19)
    End If
    wrdApp.ActiveDocument.FormFields(arrayOne(19)).Result = range2(20)
    '21-30
    wrdApp.ActiveDocument.FormFields(arrayOne(20)).Result = range2(21)
    wrdApp.ActiveDocument.FormFields(arrayOne(21)).Result = range2(22)
    wrdApp.ActiveDocument.FormFields(arrayOne(22)).Result = range2(23)
    wrdApp.ActiveDocument.FormFields(arrayOne(23)).Result = range2(24)
    wrdApp.ActiveDocument.FormFields(arrayOne(24)).Result = range2(25)
    wrdApp.ActiveDocument.FormFields(arrayOne(25)).Result = range2(26)
    wrdApp.ActiveDocument.FormFields(arrayOne(26)).Result = range2(27)
    wrdApp.ActiveDocument.FormFields(arrayOne(27)).Result = range2(28)
    wrdApp.ActiveDocument.FormFields(arrayOne(28)).Result = range2(29)
    wrdApp.ActiveDocument.FormFields(arrayOne(29)).Result = range2(30)
'    '31-40
    wrdApp.ActiveDocument.FormFields(arrayOne(30)).Result = range2(31)
    wrdApp.ActiveDocument.FormFields(arrayOne(31)).Result = range2(32)
    wrdApp.ActiveDocument.FormFields(arrayOne(32)).Result = range2(33)
    wrdApp.ActiveDocument.FormFields(arrayOne(33)).Result = range2(34)
    wrdApp.ActiveDocument.FormFields(arrayOne(34)).Result = range2(35)
    wrdApp.ActiveDocument.FormFields(arrayOne(35)).Result = range2(36)
    wrdApp.ActiveDocument.FormFields(arrayOne(36)).Result = range2(37)
    wrdApp.ActiveDocument.FormFields(arrayOne(37)).Result = range2(38)
    wrdApp.ActiveDocument.FormFields(arrayOne(38)).Result = range2(39)
    wrdApp.ActiveDocument.FormFields(arrayOne(39)).Result = range2(40)
'    '41-50
    wrdApp.ActiveDocument.FormFields(arrayOne(40)).Result = range2(41)
    wrdApp.ActiveDocument.FormFields(arrayOne(41)).Result = range2(42)
    wrdApp.ActiveDocument.FormFields(arrayOne(42)).Result = range2(43)
    wrdApp.ActiveDocument.FormFields(arrayOne(43)).Result = range2(44)
    wrdApp.ActiveDocument.FormFields(arrayOne(44)).Result = range2(45)
    wrdApp.ActiveDocument.FormFields(arrayOne(45)).Result = range2(46)
    wrdApp.ActiveDocument.FormFields(arrayOne(46)).Result = range2(47)
    wrdApp.ActiveDocument.FormFields(arrayOne(47)).Result = range2(48)
    wrdApp.ActiveDocument.FormFields(arrayOne(48)).Result = range2(49)
    wrdApp.ActiveDocument.FormFields(arrayOne(49)).Result = range2(50)
'    '51-60
    wrdApp.ActiveDocument.FormFields(arrayOne(50)).Result = range2(51)
    wrdApp.ActiveDocument.FormFields(arrayOne(51)).Result = range2(52)
    wrdApp.ActiveDocument.FormFields(arrayOne(52)).Result = range2(53)
    wrdApp.ActiveDocument.FormFields(arrayOne(53)).Result = range2(54)
    wrdApp.ActiveDocument.FormFields(arrayOne(54)).Result = range2(55)
    wrdApp.ActiveDocument.FormFields(arrayOne(55)).Result = range2(56)
    wrdApp.ActiveDocument.FormFields(arrayOne(56)).Result = range2(57)
    wrdApp.ActiveDocument.FormFields(arrayOne(57)).Result = range2(58)
    wrdApp.ActiveDocument.FormFields(arrayOne(58)).Result = range2(59)
    wrdApp.ActiveDocument.FormFields(arrayOne(59)).Result = range2(60)
'    '61-70
    wrdApp.ActiveDocument.FormFields(arrayOne(60)).Result = range2(61)
    wrdApp.ActiveDocument.FormFields(arrayOne(61)).Result = range2(62)
    wrdApp.ActiveDocument.FormFields(arrayOne(62)).Result = range2(63)
    wrdApp.ActiveDocument.FormFields(arrayOne(63)).Result = range2(64)
    wrdApp.ActiveDocument.FormFields(arrayOne(64)).Result = range2(65)
    wrdApp.ActiveDocument.FormFields(arrayOne(65)).Result = range2(66)
    wrdApp.ActiveDocument.FormFields(arrayOne(66)).Result = range2(67)
    wrdApp.ActiveDocument.FormFields(arrayOne(67)).Result = range2(68)
    wrdApp.ActiveDocument.FormFields(arrayOne(68)).Result = range2(69)
    wrdApp.ActiveDocument.FormFields(arrayOne(69)).Result = range2(70)
'    '71-77
    wrdApp.ActiveDocument.FormFields(arrayOne(70)).Result = range2(71)
    wrdApp.ActiveDocument.FormFields(arrayOne(71)).Result = range2(72)
    wrdApp.ActiveDocument.FormFields(arrayOne(72)).Result = range2(73)
    wrdApp.ActiveDocument.FormFields(arrayOne(73)).Result = range2(74)
    wrdApp.ActiveDocument.FormFields(arrayOne(74)).Result = range2(75)
    wrdApp.ActiveDocument.FormFields(arrayOne(75)).Result = range2(76)
    wrdApp.ActiveDocument.FormFields(arrayOne(76)).Result = range2(77)
    
'3)
'
    wrdApp.ActiveDocument.SaveAs2 saveFileName, wdFormatDocumentDefault
    wrdApp.Visible = True
    wrdApp.ActiveDocument.Close SaveChanges:=False
    
    'wrdDoc.Close SaveChanges:=False
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    
'    [VBA] wrdDoc.Close ' close the document
'Set wrdDoc = Nothing
'' CLOSING the document does NOT explicitly release
'' the memory assigned for it!
'' = Nothing does explicitly release memory
'wrdApp.Quit ' close down Word
'Set wrdApp = Nothing
'' Quitting Word does NOT explicitly release the memory assigned for it
'' and that is a LOT
'' = Nothing does explicitly release memory
'[/vba]
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
        
    Worksheets("Sheet1").Activate
End Sub

Sub ImportFromApp()
 
    Dim wrdApp As Object
    Dim wrdDoc As Object
    Dim myFile As String
    Dim str As String
    Dim i As Integer
    Dim Array1() As Variant 'Used to store the list of bookmark names, gotten from the ListOfBookmarks tab
    Dim arrayName() As Variant
    Dim arrayValue() As Variant
    Dim rng1 As Range 'Contains the bookmark names
    Dim rng2 As Range 'Contains the info from the application
    Dim arrayLength As Integer
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Worksheets("ApplicationImport").Activate
    
    arrayLength = Range("C4").Value2 'Resize array
    ReDim Array1(0 To arrayLength)
    
    Set rng2 = Range("A2", "EM2") 'update this if more bookmarks are added to the file. Needs to match up with row 1
    
    Set rng1 = Range("A1", Range("A1").End(xlToRight)) 'This loop puts the range into the array
    For i = 1 To arrayLength
        Array1(i) = rng1(i)
    Next i

    'Opens the application file
    myFile = Range("FileName")
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = False
    Set wrdDoc = wrdApp.Documents.Open(myFile)

    For i = 1 To arrayLength 'Uncomment below to list results in column C for debugging
        str = wrdApp.ActiveDocument.FormFields(Array1(i)).Result
        rng2(i) = str
    Next i

    wrdDoc.Close SaveChanges:=False
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

    arrayName = rng1.Value2
    arrayValue = rng2.Value2

    For i = 1 To arrayLength
        Range(arrayName(1, i)) = arrayValue(1, i)
    Next i

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Calculate
    Call UpdateAfterEvents
        
    Worksheets("Sheet1").Activate
    
End Sub
