Attribute VB_Name = "Module2"
Option Explicit

Sub ColorGreen()

    With Range("RangeToColor").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12648384
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub ColorRed()

    With Range("RangeToColor").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12632319
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub ChangeProcessorColors()

Dim FieldsToClear As Variant
Dim x As Variant
Dim Rg1 As Range

'Add fields to change to array below
FieldsToClear = Array("Prop1Info", "Prop2Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", _
        "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info")
' FieldsToClear = Array("B1Info", "B2Info", "B3Info", "LoanInfo", "LendingInfo", "TLTAInfo", _
'        "Prop1Info", "Prop2Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", _
'        "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "Prop16Info", "Prop17Info", "Prop18Info", _
'        "Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info", "MiscInfo", "AmountTotalsInfo", "Notes")

Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.EnableEvents = False

For Each x In FieldsToClear 'Changes all the cells in the array above to a blue color
       With Range(x).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 33
        .TintAndShade = 0
        .PatternTintAndShade = 0
       End With
Next x

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub ChangeLOColors()

Dim FieldsToClear As Variant
Dim x As Variant
Dim Rg1 As Range

 FieldsToClear = Array("B1Info", "B2Info", "B3Info", "LoanInfo", "LendingInfo", "TLTAInfo", _
        "Prop1Info", "Prop2Info", "Prop4Info", "Prop5Info", "Prop6Info", "Prop7Info", "Prop8Info", "Prop9Info", _
        "Prop10Info", "Prop11Info", "Prop12Info", "Prop13Info", "Prop14Info", "Prop15Info", "MiscInfo", "AmountTotalsInfo", "Notes")
        '"Prop16Info", "Prop17Info", "Prop18Info","Prop19Info", "Prop20Info", "Prop21Info", "Prop22Info", "Prop23Info", "Prop24Info", "Prop25Info",
Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.EnableEvents = False

For Each x In FieldsToClear 'Sets them all to the default grey color
       With Range(x).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
       End With
Next x

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
