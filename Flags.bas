Attribute VB_Name = "Flags"
Option Explicit

Sub UpdateFlags() 'master flag method. All other methods should call this one.

Call AoIFlag

End Sub

Sub AoIFlag() 'Affadavit of Identity Flag

Dim flag1 As Integer
Dim flag2 As Integer
Dim flag3 As Integer

If Range("Borrower1FKA") = 0 And Range("Borrower1AKA1") = 0 Then
    flag1 = 0
Else
    flag1 = 1
End If
Range("HideAoI1") = flag1

If Range("Borrower2FKA") = 0 And Range("Borrower2AKA1") = 0 Then
    flag2 = 0
Else
    flag2 = 1
End If
Range("HideAoI2") = flag2

If Range("Borrower3FKA") = 0 And Range("Borrower3AKA1") = 0 Then
    flag3 = 0
Else
    flag3 = 1
End If
Range("HideAoI3") = flag3

End Sub
