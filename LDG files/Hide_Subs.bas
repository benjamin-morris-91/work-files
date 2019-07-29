Attribute VB_Name = "Hide_Subs"
Option Explicit

Sub HideBorrowers() 'Hides optional borrowers until needed

Dim WS1 As Worksheet
Set WS1 = Worksheets("Sheet1")

If Range("NumberOfBorrowers").Value = 3 Then
        WS1.Rows("21:39").EntireRow.Hidden = False
    Else
        Dim HideRows1 As Integer
        HideRows1 = (21 + (9 * (Range("NumberOfBorrowers").Value - 1)))
        
        WS1.Rows("21:39").EntireRow.Hidden = False
        WS1.Rows(HideRows1 & ":39").EntireRow.Hidden = True
    End If

Call prefillBorrowerInfo

End Sub
Sub updateHomeStead()

If Range("LoanType") = "Homestead - New Loan" Then
    Range("LoanTypeChoice") = "Yes"
Else
    Range("LoanTypeChoice") = "No"
End If

End Sub
Sub IndividualLoan()
    Range("Borrower1Marital") = ""
    Range("Borrower2Marital") = ""
    Range("Borrower3Marital") = ""
    Range("Borrower1DOB") = ""
    Range("Borrower2DOB") = ""
    Range("Borrower3DOB") = ""

End Sub

Sub CorporateLoan()
    Range("Borrower1Marital") = "N/A"
    Range("Borrower2Marital") = "N/A"
    Range("Borrower3Marital") = "N/A"
    Range("Borrower1DOB") = "N/A"
    Range("Borrower2DOB") = "N/A"
    Range("Borrower3DOB") = "N/A"

End Sub

Sub LLCLoan()
    Range("Borrower1Marital") = "N/A"
    Range("Borrower2Marital") = "N/A"
    Range("Borrower3Marital") = "N/A"
    Range("Borrower1DOB") = "N/A"
    Range("Borrower2DOB") = "N/A"
    Range("Borrower3DOB") = "N/A"

End Sub

Sub CopyAddress() 'Copies B1 physical address to B2 and B3

Range("Borrower2AddressStreet") = Range("Borrower1AddressStreet")
Range("Borrower2AddressCity") = Range("Borrower1AddressCity")
Range("Borrower2AddressState") = Range("Borrower1AddressState")
Range("Borrower2AddressZIP") = Range("Borrower1AddressZIP")
Range("Borrower3AddressStreet") = Range("Borrower1AddressStreet")
Range("Borrower3AddressCity") = Range("Borrower1AddressCity")
Range("Borrower3AddressState") = Range("Borrower1AddressState")
Range("Borrower3AddressZIP") = Range("Borrower1AddressZIP")

Range("B2MAddress") = Range("MailingStreet") 'Copies B1 mailing address to B2 and B3
Range("B2MCity") = Range("MailingCity")
Range("B2MState") = Range("MailingState")
Range("B2MZIP") = Range("MailingZIP")
Range("B3MAddress") = Range("MailingStreet")
Range("B3MCity") = Range("MailingCity")
Range("B3MState") = Range("MailingState")
Range("B3MZIP") = Range("MailingZIP")

End Sub

Sub HideProperties() 'Hides additional Property Info unless needed

Dim WS1 As Worksheet
Set WS1 = Worksheets("Sheet1")

    If Range("NumberOfProperties").Value = 25 Then
        WS1.Rows("82:345").EntireRow.Hidden = False
    Else
        Dim HideRows As Integer
        HideRows = (82 + (11 * (Range("NumberOfProperties").Value - 1))) 'Below calculation figures out the row to start hiding
        WS1.Rows("82:345").EntireRow.Hidden = False
        WS1.Rows(HideRows & ":345").EntireRow.Hidden = True
    End If
    
    Call prefillPropAns 'prefills property info after user selects the number of borrowers

End Sub

Sub HideActiveXBoxes() 'Hides the ActiveXBoxes based on the number of properties selected.

Dim x As Integer
Dim str As Variant
Dim numberOfProperties As Integer

numberOfProperties = Range("NumberofProperties")


Select Case numberOfProperties
    Case Is = 1
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = False
        Worksheets("Sheet1").CommandButton3.Visible = False
        Worksheets("Sheet1").CommandButton4.Visible = False
        Worksheets("Sheet1").CommandButton5.Visible = False
        Worksheets("Sheet1").CommandButton6.Visible = False
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 2
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = False
        Worksheets("Sheet1").CommandButton4.Visible = False
        Worksheets("Sheet1").CommandButton5.Visible = False
        Worksheets("Sheet1").CommandButton6.Visible = False
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 3
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = False
        Worksheets("Sheet1").CommandButton5.Visible = False
        Worksheets("Sheet1").CommandButton6.Visible = False
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 4
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = False
        Worksheets("Sheet1").CommandButton6.Visible = False
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 5
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = False
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False
    
    Case Is = 6
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = False
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 7
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = False
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 8
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = False
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 9
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = False
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 10
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = False
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 11
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = True
        Worksheets("Sheet1").CommandButton12.Visible = False
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 12
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = True
        Worksheets("Sheet1").CommandButton12.Visible = True
        Worksheets("Sheet1").CommandButton13.Visible = False
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 13
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = True
        Worksheets("Sheet1").CommandButton12.Visible = True
        Worksheets("Sheet1").CommandButton13.Visible = True
        Worksheets("Sheet1").CommandButton14.Visible = False
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Is = 14
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = True
        Worksheets("Sheet1").CommandButton12.Visible = True
        Worksheets("Sheet1").CommandButton13.Visible = True
        Worksheets("Sheet1").CommandButton14.Visible = True
        Worksheets("Sheet1").CommandButton15.Visible = False

    Case Else '15 properties, do nothing
        Worksheets("Sheet1").CommandButton1.Visible = True
        Worksheets("Sheet1").CommandButton2.Visible = True
        Worksheets("Sheet1").CommandButton3.Visible = True
        Worksheets("Sheet1").CommandButton4.Visible = True
        Worksheets("Sheet1").CommandButton5.Visible = True
        Worksheets("Sheet1").CommandButton6.Visible = True
        Worksheets("Sheet1").CommandButton7.Visible = True
        Worksheets("Sheet1").CommandButton8.Visible = True
        Worksheets("Sheet1").CommandButton9.Visible = True
        Worksheets("Sheet1").CommandButton10.Visible = True
        Worksheets("Sheet1").CommandButton11.Visible = True
        Worksheets("Sheet1").CommandButton12.Visible = True
        Worksheets("Sheet1").CommandButton13.Visible = True
        Worksheets("Sheet1").CommandButton14.Visible = True
        Worksheets("Sheet1").CommandButton15.Visible = True
End Select




End Sub

'Sub HideLoanType() 'Hides the File Number Row if not needed
'
'Dim WS1 As Worksheet
'Set WS1 = Worksheets("Sheet1")
'
'    If Range("LoanType").Value = "Investment Residential" Then
'        WS1.Rows("6").EntireRow.Hidden = False
'    Else
'        WS1.Rows("6").EntireRow.Hidden = True
'    End If
'End Sub
