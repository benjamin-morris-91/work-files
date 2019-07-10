Attribute VB_Name = "UserFormControl"
Option Explicit
'Contains the Userforms

Sub ShowLoadLoanUF()

'This runs before showing the new screen.

Application.ScreenUpdating = False
Call LoadFromDatabase
Worksheets("Sheet1").Activate
Application.ScreenUpdating = True

Load_Loan.Show

End Sub

Sub ShowLegalBox()

LegalBoxDriver.Show

End Sub

Sub ShowInstructionBox()

DisplayInstructions.Show

End Sub
