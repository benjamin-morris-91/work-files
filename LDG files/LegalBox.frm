VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LegalBox 
   Caption         =   "Legal Description"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "LegalBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LegalBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
'Saving Legal
Dim number As String
Dim LegalString1 As String


LegalString1 = "Prop" & LegalBoxDriver.ComboBox1.Value & "Legal"

Range(LegalString1) = LegalBox.LegalTextBox
Unload LegalBox

End Sub

Private Sub CommandButton2_Click()
'Canceling Legal
Unload LegalBox

End Sub

'
'Private Sub UserForm_Initialize()
'    PickPropertyCB.AddItem "1"
'    PickPropertyCB.AddItem "2"
'    PickPropertyCB.AddItem "3"
'    PickPropertyCB.AddItem "4"
'    PickPropertyCB.AddItem "5"
'    PickPropertyCB.AddItem "6"
'    PickPropertyCB.AddItem "7"
'    PickPropertyCB.AddItem "8"
'    PickPropertyCB.AddItem "9"
'    PickPropertyCB.AddItem "10"
'    PickPropertyCB.AddItem "11"
'    PickPropertyCB.AddItem "12"
'    PickPropertyCB.AddItem "13"
'    PickPropertyCB.AddItem "14"
'    PickPropertyCB.AddItem "15"
'
'End Sub
Private Sub LegalTextBox_Change()

End Sub
