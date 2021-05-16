VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnicodeFromVBADisplay 
   Caption         =   "Display for Hardcoded Unicode in VBA"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   OleObjectBlob   =   "UnicodeFromVBADisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnicodeFromVBADisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProgrammingTools")
Option Explicit

Private Sub btnOK_Click()
    Me.Hide
    Unload Me
    Application.VBE.ActiveWindow.SetFocus
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
    Application.VBE.ActiveCodePane.Show
End Sub
