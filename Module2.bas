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
