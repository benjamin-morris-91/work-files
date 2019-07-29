VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeUserColor 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ChangeUserColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeUserColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButton1_Click()

If ChangeUserColor.OptionButton1 = True Then
    Call ChangeLOColors
End If

If ChangeUserColor.OptionButton2 = True Then
    Call ChangeProcessorColors
End If

Unload ChangeUserColor

End Sub
