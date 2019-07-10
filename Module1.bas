Attribute VB_Name = "Module1"
Sub CalculateRecordingCosts()

Dim county As String
Dim number As Integer
Dim i As Integer
Dim j As Integer
Dim numberOfEntities As Integer
Dim Entity As String
Dim recordingTotal As Integer

'county = InputBox("Enter the county")
'number = InputBox("Enter the length of tax terms")
'Worksheets("RecordingCosts").Range("F1").Value = county
'Worksheets("RecordingCosts").Range("F2").Value = number

For i = 1 To Range("NumberOfProperties")
    numberOfEntities = 0
    For j = 1 To 4
        Entity = "Prop" & i & "TaxingEntity" & j
        If Range(Entity).Value <> "0" Then
            numberOfEntities = numberOfEntities + 1
        End If
    Next j
    Worksheets("RecordingCosts").Range("F4") = numberOfEntities
    recordingTotal = recordingTotal + Worksheets("RecordingCosts").Range("F10")
    
Next i
MsgBox recordingTotal

End Sub

Sub AddB1FKA()

Range("CombinedB1FKA") = Range("Borrower1Name") & " fka " & Range("Borrower1FKA")

End Sub

Sub AddB2FKA()

Range("CombinedB2FKA") = Range("Borrower2Name") & " fka " & Range("Borrower2FKA")

End Sub

Sub AddB3FKA()

Range("CombinedB3FKA") = Range("Borrower3Name") & " fka " & Range("Borrower3FKA")

End Sub

Sub LegalLoop()
'set up for loop to pass through the legal description as a function, return false if any constraints fail, which fails this sub.
Dim str As String

Dim a As Boolean
Dim x As Integer

For x = 1 To Range("NumberOfProperties")
    str = Range("Prop" & x & "Legal")
    a = LegalWork(str)
    If a = False Then
        Exit For
    End If
Next x
Range("LegalCriteria") = a

End Sub

Public Function LegalWork(str As String) As Boolean
'***************************************************
'Checks to ensure users abides by Legal Description restraints.
'Restraints: Legal 1 is <=3 lines and <=170 characters total.
'Does NOT check if a users enters more than 56 characters on one line. Couldn't figure it out...
'***************************************************
Dim legalLength As Integer
Dim lineLength As Integer

'legalLength = Len(Range("Prop1Legal"))
legalLength = Len(str)

'lineLength = Worksheets("DropdownInfo").Range("B146")
lineLength = Len(str) - Len(WorksheetFunction.Substitute(str, Chr(10), "")) + 1
If legalLength > 170 Then
    MsgBox "The length of the lines exceed the max allowed. Please make sure to have less than 170 characters and use the Exhabit A box for anything that exceeds 170 characters"
    LegalWork = False
    Exit Function
ElseIf lineLength > 3 Then
    MsgBox "The number of lines used exceeds the max allowed. Please make sure to place any lines after 3 in the exhibit A box"
    LegalWork = False
    Exit Function
Else
    LegalWork = True
End If

End Function

Sub LogIssue() 'Adds the ability to log an issue with the program. This gets saved to the database file.

Dim Wb2 As Workbook                     'Database Doc
Dim newDatabase As Worksheet            'copying from main file
Dim UpdateLog As Worksheet     'pasting into database file
Dim inputSelection As Variant
Dim str As String
Dim nextSpot As Integer
Dim strA As String
Dim strB As String

inputSelection = InputBox("Please enter the problem you encountered.")
    If inputSelection <> "" Then 'Process of adding the entry to the log
    
    ' ********************
    ' Only emailing issue. Currently error in sending it to database file. Keeps overwritting the same cell...
    ' ********************
    
    
        Application.ScreenUpdating = False
        Call assignFileNames
        Application.ScreenUpdating = False

        Set Wb2 = Workbooks.Open(databaseFile) 'This is the database file
        Set UpdateLog = Worksheets("LogIssues")
        
        Worksheets("LogIssues").Activate
        nextSpot = Application.WorksheetFunction.CountA(Columns(1)) + 1

        strA = "A" & nextSpot
        strB = "F" & nextSpot
        str = Now() & " - " & Application.UserName
        Worksheets("LogIssues").Range(strA) = str
        Worksheets("LogIssues").Range(strB) = inputSelection
        Worksheets("Sheet1").Activate
        
        Wb2.Close SaveChanges:=True
        Set UpdateDatabaseFile = Nothing
        Application.ScreenUpdating = True
        
        EmailIssues (inputSelection) 'Emails issues
        MsgBox "Your issue has been recorded."
    End If

End Sub

Function EmailIssues(strComplaint As String) 'Receives a str from LogIssue() that emails issue to me (ben@hunterkelsey.com)

Dim strTime As String
Dim strPerson As String

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")

Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)

strTime = "Issue Logged at " & Now
strPerson = "Issue logged by " & Application.UserName

olMail.To = "ben@hunterkelsey.com"
olMail.Subject = strTime
olMail.Body = strComplaint & Chr(10) & Chr(10) & strPerson
olMail.Send

End Function

Sub splitLegals() 'Separates the legal description depending on the size.

Dim currentLegal As String
Dim strA As String
Dim strB As String
Dim i As Integer
Dim length As Integer

Worksheets("DropdownInfo").Range("M48:N72").ClearContents

For i = 1 To Range("NumberofProperties")
    currentLegal = "Prop" & i & "Legal"
    strA = "Prop" & i & "LegalA"
    strB = "Prop" & i & "LegalB"
    length = Len(Range(currentLegal))
    
    Range(strA) = ""
    Range(strB) = ""
    
    If length > 300 Then
        Range(strA) = "See Exhibit A"
        Range(strB) = Range(currentLegal)
    Else
        Range(strA) = Range(currentLegal)
    End If
Next i

End Sub





