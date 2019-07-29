Attribute VB_Name = "Buttons"
Option Explicit

Sub CopyB1AddressToProperties()
' Copies the address from B1 to each of the properties
' Will exit sub if user clicks the OK, Cancel, or the X in the input box
' Program will loop until user enters a valid input (1, 2, or 3) or exits by entering the above conditions.

Dim strAddress As String
Dim strCity As String
Dim strState As String
Dim strZIP As String
Dim BorrowerAddress As String
Dim BorrowerCity As String
Dim BorrowerState As String
Dim BorrowerZIP As String
Dim i As Integer
Dim j As Integer
Dim inputSelection As Variant
Dim validCheck As Boolean

inputSelection = InputBox("Which Borrower would you like to copy the address from? 1, 2 or 3?")

If inputSelection = "" Then
    Exit Sub
ElseIf inputSelection >= 1 And inputSelection <= 3 Then
    validCheck = True
Else
    validCheck = False
End If

If validCheck = True Then
    j = inputSelection 'Necessary to remove any potential leading zeros
    If j = 1 Then
        BorrowerAddress = "Borrower1PAddress"
        BorrowerCity = "Borrower" & j & "AddressCity"
        BorrowerState = "Borrower" & j & "AddressState"
        BorrowerZIP = "Borrower" & j & "AddressZIP"
    Else
        BorrowerAddress = "Borrower" & j & "AddressStreet"
        BorrowerCity = "Borrower" & j & "AddressCity"
        BorrowerState = "Borrower" & j & "AddressState"
        BorrowerZIP = "Borrower" & j & "AddressZIP"
    End If

    For i = 1 To Range("NumberOfProperties")
        strAddress = "Prop" & i & "Address"
        strCity = "Prop" & i & "City"
        strState = "Prop" & i & "State"
        strZIP = "Prop" & i & "ZIP"

        Range(strAddress) = Range(BorrowerAddress)
        Range(strCity).Value = Range(BorrowerCity)
        Range(strState).Value = Range(BorrowerState)
        Range(strZIP).Value = Range(BorrowerZIP)
    Next i
Else
    MsgBox "You must enter either a 1, 2, or 3!"
    Run "CopyB1AddressToProperties"

End If

End Sub


