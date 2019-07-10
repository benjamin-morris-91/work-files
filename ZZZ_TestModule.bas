Attribute VB_Name = "ZZZ_TestModule"
Option Explicit
'******************* Test Module for testing new features ************************

Public FSO As New FileSystemObject 'Creating a FileSystemObject

Sub TestingDiskSpace()

    Dim drv As Drive
    Dim Space As Double
    Set drv = FSO.GetDrive("C:") ' Creating the the Drive object
    Space = drv.FreeSpace

    Space = Space / 1073741824 'converting bytes to GB
    Space = WorksheetFunction.Round(Space, 2) ' Rounding
    MsgBox "C: has free space = " & Space & " GB"

End Sub


Sub TestingCopy()
    
    Call assignFileNames
    Application.ScreenUpdating = False

    LoanDocDB.Range("A1", Range("A2").End(xlToRight)).Copy
    'LoanDocDB.Range("A20").Select
    
    LoanDocDB.Range("A20").PasteSpecial xlPasteValues  'PasteSpecial xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    LoanDocDB.Range("A20").PasteSpecial xlPasteFormats


    Application.CutCopyMode = False

End Sub

Sub TestingCopyingFormats()

Dim i As Double

Worksheets("Driver").Range("J20") = Range("UniqueAmount1")

    Worksheets("Driver").Range("J20:J20").Select
    With Selection
        .NumberFormat = "$#,##0.00"
        .Value = .Value
    End With
    
i = Range("J20")
MsgBox i


End Sub

Sub TestChangingColors()
'
' TestChangingColors Macro
'

'
    With Range("A1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Range("A1").Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
    End With
    
End Sub

Sub TestingErrorStuff() 'Check to see if something is a number or not.

Dim inputSelection As Variant

inputSelection = InputBox("Please enter the problem you encountered.")

If IsNumeric(inputSelection) = True Then
    MsgBox "It is a number"
Else
    MsgBox "It is not a number"
End If


End Sub
