VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnicodeToVBAPrompt 
   Caption         =   "Unicode Encoder for VBA"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690.001
   OleObjectBlob   =   "UnicodeToVBAPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnicodeToVBAPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProgrammingTools")
Option Explicit

Private Sub btnCancel_Click()
    Hide
    Unload Me
End Sub

Private Sub btnInsert_Click()
    Hide
    
    Dim vbaCode As String
    vbaCode = ProgrammingTools.ConvertUnicodeTextToVBACode(tbxInput.Value)
    
    Unload Me
    
    Dim pane As CodePane
    Dim codeMod As CodeModule
    Set pane = Application.VBE.ActiveCodePane
    Set codeMod = pane.CodeModule
    
    Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    pane.GetSelection startLine, startCol, endLine, endCol
    Dim textBeforeSelection As String, textAfterSelection As String
    textBeforeSelection = Left$(codeMod.Lines(startLine, 1), startCol - 1)
    textAfterSelection = Mid$(codeMod.Lines(endLine, 1), endCol + 1)
    codeMod.DeleteLines startLine, endLine - startLine + 1
    codeMod.InsertLines startLine, textBeforeSelection & vbaCode & textAfterSelection
    
    Application.VBE.ActiveCodePane.Show
End Sub




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
    Application.VBE.ActiveCodePane.Show
End Sub
