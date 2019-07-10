Attribute VB_Name = "References"
Option Explicit

Sub AddReferenceGUID() 'Adds a reference using the GUID
'   Credit to Siddharth Rout and Ken Puls
'   https://stackoverflow.com/questions/9879825/how-to-add-a-reference-programmatically
'   Preconditions:
'   1) Macros are enabled
'   2) In security settings "Trust Access To Visual Basic Project" is checked
'   3) Manually set a reference to 'Microsoft Visual Basic for Applications Extensibility" obejct. Tools -> Reference
    
    Dim strGUID As String
    Dim theRef As Variant
    Dim i As Long
 
    strGUID = "{00062FFF-0000-0000-C000-000000000046}" 'Update this to add different reference.

    On Error Resume Next
    
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1 'Remove any missing references
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.IsBroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i
    
    Err.Clear 'Clear any errors so that error trapping for GUID additions can be evaluated

    'Adds the reference
    ThisWorkbook.VBProject.References.AddFromGuid _
    GUID:=strGUID, Major:=1, Minor:=0

    Select Case Err.number 'If an error was encountered, inform the user
    Case Is = 32813 'Reference already in use.  No action necessary
    Case Is = vbNullString 'Reference added without issue
    Case Else 'An unknown error was encountered, so alert the user
        MsgBox "A problem was encountered trying to" & vbNewLine _
        & "add or remove a reference in this file" & vbNewLine & "Please check the " _
        & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0
    
End Sub

Sub ListReferencePaths() 'Run macro to list all GUID in the immediate window. Plug these into AddReferenceGUID() to add a reference using vba code.
'   Credit to Chad Crowe
'   https://stackoverflow.com/questions/9879825/how-to-add-a-reference-programmatically

On Error Resume Next
Dim i As Long

Debug.Print "Reference name" & " | " & "Full path to reference" & " | " & "Reference GUID"

For i = 1 To ThisWorkbook.VBProject.References.Count
  With ThisWorkbook.VBProject.References(i)
    Debug.Print .Name & " | " & .FullPath & " | " & .GUID
  End With
Next i
On Error GoTo 0
End Sub
