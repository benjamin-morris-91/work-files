VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveUF 
   Caption         =   "Override"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "SaveUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

Range("SaveChoice") = 1
Unload SaveUF

End Sub

Private Sub CommandButton2_Click()

Range("SaveChoice") = 2
Unload SaveUF

End Sub
