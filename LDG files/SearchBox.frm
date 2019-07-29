VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchBox 
   Caption         =   "UserForm1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   OleObjectBlob   =   "SearchBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoadSelected_Click()

Dim RowNum As Long
Dim ListBoxRow As Long

If IsNull(listResults) Then
    MsgBox "Nothing selected"
Else
    If Worksheets("TempDatabase").Range("J1").Value = 1 Then 'Searching by Loan Number
        Worksheets("TempDatabase").Range("J2").Value = listResults.Column(0)
        Worksheets("TempDatabase").Range("J3").Value = listResults.Column(1)
    Else 'Searching by Name
        Worksheets("TempDatabase").Range("J2").Value = listResults.Column(1)
        Worksheets("TempDatabase").Range("J3").Value = listResults.Column(0)
    End If
    
    Worksheets("TempDatabase").Range("J4").Value = Worksheets("TempDatabase").Range("J2") & " - " & Worksheets("TempDatabase").Range("J3")
    Range("LNToLoad") = Worksheets("TempDatabase").Range("J4")
    
    Unload SearchBox
    Call LoadLoan
    
End If

End Sub

Private Sub cmdSearchNumber_Click() 'Search Loan Number button
'Preconditions: Should have already loaded the database into the "NewDatabase" tab

Dim RowNum As Long
Dim SearchRow As Long

RowNum = 4
SearchRow = 2

Worksheets("TempDatabase").Range("A2:H1000").ClearContents
Worksheets("NewDatabase").Activate

Do Until Cells(RowNum, 1).Value = ""
    If InStr(1, Cells(RowNum, 1).Value, TextBox1.Value, vbTextCompare) > 0 Then
        Worksheets("TempDatabase").Cells(SearchRow, 1).Value = Cells(RowNum, 1).Value
        Worksheets("TempDatabase").Cells(SearchRow, 2).Value = Cells(RowNum, 67).Value
        Worksheets("TempDatabase").Cells(SearchRow, 3).Value = Cells(RowNum, 52).Value
        SearchRow = SearchRow + 1
    End If
    RowNum = RowNum + 1
Loop

If SearchRow = 2 Then
    Worksheets("Sheet1").Activate
    MsgBox "No products were found that match your search criteria."
    Exit Sub
End If

Worksheets("TempDatabase").Range("J1").Value = 1

listResults.RowSource = "SearchResults"
Worksheets("Sheet1").Activate

End Sub

Private Sub cmdSearchName_Click() 'Search Name button
'Preconditions: Should have already loaded the database into the "NewDatabase" tab

Dim RowNum As Long
Dim SearchRow As Long

RowNum = 4
SearchRow = 2

Worksheets("TempDatabase").Range("A2:H1000").ClearContents
Worksheets("NewDatabase").Activate

Do Until Cells(RowNum, 67).Value = ""   'Searches column 67, which is B1Name currently.
    If InStr(1, Cells(RowNum, 67).Value, TextBox1.Value, vbTextCompare) > 0 Then
        Worksheets("TempDatabase").Cells(SearchRow, 6).Value = Cells(RowNum, 67).Value  'This correlates to B1Name
        Worksheets("TempDatabase").Cells(SearchRow, 7).Value = Cells(RowNum, 1).Value   'This correlates to Loan Number
        Worksheets("TempDatabase").Cells(SearchRow, 8).Value = Cells(RowNum, 52).Value  'This correlates to Application Date
        SearchRow = SearchRow + 1
    End If
    RowNum = RowNum + 1
Loop

If SearchRow = 2 Then
    Worksheets("Sheet1").Activate
    MsgBox "No products were found that match your search criteria."
    Exit Sub
End If

Worksheets("TempDatabase").Range("J1").Value = 2

listResults.RowSource = "SearchResults2"
Worksheets("Sheet1").Activate

End Sub

Private Sub UserForm_Initialize()

listResults.SetFocus
Worksheets("TempDatabase").Range("A2:H1000").ClearContents

End Sub

