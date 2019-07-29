VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LegalBoxDriver 
   Caption         =   "Add/Edit Legal"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "LegalBoxDriver.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LegalBoxDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButton1_Click()

'*************************************************
'Does not handle if user entered a string. Crashes.
'*************************************************

Dim number As String
Dim LegalString As String
Dim flag1 As Boolean

number = LegalBoxDriver.ComboBox1
flag1 = True 'flag to decide if LegalBox should be opened.

If number = "" Then
    MsgBox "You need to pick a property"
    Unload LegalBoxDriver
    number = 0
    flag1 = False
ElseIf number > 15 Or number < 1 Then
    MsgBox "You entered an invalid number"
    Unload LegalBoxDriver
    flag1 = False
End If

LegalString = "Prop" & number & "Legal"


If flag1 = True Then
    If AddLegal.Value = True Then 'If Add Legal Option Button is clicked
        LegalBox.Show
        Unload LegalBoxDriver
    
    ElseIf EditLegal.Value = True Then 'If Edit Legal Option Button is clicked
        LegalBox.LegalTextBox.Value = Range(LegalString)
        LegalBox.Show
        Unload LegalBoxDriver

    Else
        MsgBox "You need to pick whether to add or edit the legal description"
        Unload LegalBoxDriver
    End If
End If

End Sub



Private Sub UserForm_Initialize()

    ComboBox1.AddItem "1"
    ComboBox1.AddItem "2"
    ComboBox1.AddItem "3"
    ComboBox1.AddItem "4"
    ComboBox1.AddItem "5"
    ComboBox1.AddItem "6"
    ComboBox1.AddItem "7"
    ComboBox1.AddItem "8"
    ComboBox1.AddItem "9"
    ComboBox1.AddItem "10"
    ComboBox1.AddItem "11"
    ComboBox1.AddItem "12"
    ComboBox1.AddItem "13"
    ComboBox1.AddItem "14"
    ComboBox1.AddItem "15"

End Sub
