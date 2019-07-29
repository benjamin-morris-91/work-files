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
    Dim ArrayLength As Integer
    
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Worksheets("ApplicationImport").Activate
    
    ArrayLength = Range("C4").Value2 'Resize array
    ReDim Array1(0 To ArrayLength)
    
    Set rng2 = Range("A2", "EM2") 'update this if more bookmarks are added to the file. Needs to match up with row 1
    
    Set rng1 = Range("A1", Range("A1").End(xlToRight)) 'This loop puts the range into the array
    For i = 1 To ArrayLength
        Array1(i) = rng1(i)
    Next i

    'Opens the application file
    myFile = Range("FileName")
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = False
    Set wrdDoc = wrdApp.Documents.Open(myFile)

    For i = 1 To ArrayLength 'Uncomment below to list results in column C for debugging
        str = wrdApp.ActiveDocument.FormFields(Array1(i)).Result
        rng2(i) = str
    Next i

    wrdDoc.Close SaveChanges:=False
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

    arrayName = rng1.Value2
    arrayValue = rng2.Value2

    For i = 1 To ArrayLength
        Range(arrayName(1, i)) = arrayValue(1, i)
    Next i

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Calculate
    Call UpdateAfterEvents
        
    Worksheets("Sheet1").Activate
    
End Sub
