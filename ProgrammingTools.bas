Attribute VB_Name = "ProgrammingTools"
Attribute VB_Description = "Some subroutines and functions to make it easier to deal with Unicode in VBA."
'@ModuleDescription "Some subroutines and functions to make it easier to deal with Unicode in VBA."
'@Folder("VBAProgrammingTools")
'#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
'#Else
'    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
'#End If
'
'#If VBA7 Then
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
'#Else
'    Private Declare Function MessageBoxW Lib "user32" () As Long
'#End If

Private Enum UnicodeToVBACodeConverterState
    Start
    InQuote
    NotInQuote
End Enum

Private Const VBA_MAX_LINE_LENGTH As Long = 512 'Actually 1024 but 512 character lines are already crazy long

'Adapted from post by John_w at https://www.mrexcel.com/board/threads/vba-display-foreign-character-code.1142510/post-5536387
'@Description "Unicode-friendly version of MsgBox."
Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft PowerPoint") As VbMsgBoxResult
    Prompt = Prompt & vbNullChar                 'Add null terminators
    Title = Title & vbNullChar
    MsgBoxW = MessageBoxW(GetActiveWindow(), StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

'Public Sub ShowCharForCodePoint()
'    Dim selectedCode As String
'    selectedCode = GetSelectedVBACode()
'    MsgBoxW ChrW$(CLng(selectedCode))
'End Sub

' &H3001

'@EntryPoint
'@Description "Opens a window to display Unicode text corresponding to the currently selected code in the VB Editor. Useful for checking what a long sequence of calls to ChrW$() represent."
Public Sub DisplayUnicodeFromVBACode()
Attribute DisplayUnicodeFromVBACode.VB_Description = "Opens a window to display Unicode text corresponding to the currently selected code in the VB Editor. Useful for checking what a long sequence of calls to ChrW$() represent."
    Dim selectedCode As String
    selectedCode = GetSelectedVBACode()
    selectedCode = ConvertChrWCallsToUnicode(selectedCode)
    
'    'variables used in both loops to catch hex codepoints and decimal codepoints
'    Dim i As Long
'    Dim codepoint As Long
'
'    'replace code with hardcoded hex codepoints
'    Dim codepointInChrWFinder As New RegExp
'    codepointInChrWFinder.pattern = "ChrW\$?\((&H[0-9A-Fa-f]+)\)"
'    codepointInChrWFinder.Global = True
'    Dim hexInChrWs As MatchCollection
'    Set hexInChrWs = codepointInChrWFinder.Execute(selectedCode)
'
'    Dim hexMatch As Match
'    For i = hexInChrWs.Count - 1 To 0 Step -1
'        Set hexMatch = hexInChrWs.Item(i)
'        codepoint = CLng(hexMatch.SubMatches.Item(0))
'        selectedCode = Left$(selectedCode, hexMatch.FirstIndex) & """" & ChrW$(codepoint) & """" & Mid$(selectedCode, hexMatch.FirstIndex + 1 + hexMatch.Length)
'    Next i
'
'    'replace code with hardcoded hex codepoints
'    codepointInChrWFinder.pattern = "ChrW\$?\((\d+)\)"
'    Dim decInChrWs As MatchCollection
'    Set decInChrWs = codepointInChrWFinder.Execute(selectedCode)
'
'    Dim decMatch As Match
'    For i = decInChrWs.Count - 1 To 0 Step -1
'        Set decMatch = decInChrWs.Item(i)
'        codepoint = CLng(decMatch.SubMatches.Item(0))
'        selectedCode = Left$(selectedCode, decMatch.FirstIndex) & """" & ChrW$(codepoint) & """" & Mid$(selectedCode, decMatch.FirstIndex + 1 + decMatch.Length)
'    Next i
'
'    selectedCode = Replace(selectedCode, """ & """, "")
    Load UnicodeFromVBADisplay
    UnicodeFromVBADisplay.tbxUnicodeDisplay.Value = selectedCode
    UnicodeFromVBADisplay.Show vbModeless
End Sub

'@Description "Replaces ChrW() calls on decimal or hexadecimal literals inside of a string of VBA code with the characters they represent."
Public Function ConvertChrWCallsToUnicode(code As String) As String
    'variables used in both loops to catch hex codepoints and decimal codepoints
    Dim i As Long
    Dim codepoint As Long
    
    Dim output As String
    output = code
    
    'replace code with hardcoded hex codepoints
    Dim codepointInChrWFinder As New RegExp
    codepointInChrWFinder.pattern = "ChrW\$?\((&H[0-9A-Fa-f]+)\)"
    codepointInChrWFinder.Global = True
    Dim hexInChrWs As MatchCollection
    Set hexInChrWs = codepointInChrWFinder.Execute(output)
    
    Dim hexMatch As Match
    For i = hexInChrWs.Count - 1 To 0 Step -1
        Set hexMatch = hexInChrWs.Item(i)
        codepoint = CLng(hexMatch.SubMatches.Item(0))
        output = Left$(output, hexMatch.FirstIndex) & """" & ChrW$(codepoint) & """" & Mid$(output, hexMatch.FirstIndex + 1 + hexMatch.Length)
    Next i
    
    'replace code with hardcoded hex codepoints
    codepointInChrWFinder.pattern = "ChrW\$?\((\d+)\)"
    Dim decInChrWs As MatchCollection
    Set decInChrWs = codepointInChrWFinder.Execute(output)
    
    Dim decMatch As Match
    For i = decInChrWs.Count - 1 To 0 Step -1
        Set decMatch = decInChrWs.Item(i)
        codepoint = CLng(decMatch.SubMatches.Item(0))
        output = Left$(output, decMatch.FirstIndex) & """" & ChrW$(codepoint) & """" & Mid$(output, decMatch.FirstIndex + 1 + decMatch.Length)
    Next i
    
    output = Replace(output, """ & """, "")
    
    ConvertChrWCallsToUnicode = output
    Exit Function
End Function

'Test = ChrW$(&H3053) & ChrW$(&H3093) & ChrW$(&H306B) & ChrW$(&H3061) & ChrW$(&H306F) & ChrW$(&HFF01)
'ChrW$(12376) & ChrW$(12419) & ChrW$(&H3042) & ChrW$(12397) & ChrW$(65281)

'@Description "Returns a string containing the code currently selected in the VB Editor."
Public Function GetSelectedVBACode() As String
Attribute GetSelectedVBACode.VB_Description = "Returns a string containing the code currently selected in the VB Editor."
    Dim pane As CodePane
    Dim codeMod As CodeModule
    Set pane = Application.VBE.ActiveCodePane
    Set codeMod = pane.CodeModule
    
    Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    pane.GetSelection startLine, startCol, endLine, endCol
    
    Dim selectedCode As String
    selectedCode = codeMod.Lines(startLine, endLine - startLine + 1)
    If startLine = endLine Then
        selectedCode = Mid$(selectedCode, startCol, endCol - startCol)
    Else
        Dim lastLine As String
        Dim rightTrimCount As Long
        lastLine = codeMod.Lines(endLine, 1)
        rightTrimCount = Len(lastLine) - endCol
        selectedCode = Mid$(selectedCode, startCol, Len(selectedCode) - startCol - rightTrimCount)
    End If
    GetSelectedVBACode = selectedCode
End Function


'@Description "Converts a string of Unicode text to a VBA expression equivalent to that text."
Public Function ConvertUnicodeTextToVBACode(ByRef text As String) As String
Attribute ConvertUnicodeTextToVBACode.VB_Description = "Converts a string of Unicode text to a VBA expression equivalent to that text."
    Dim result As String
    Dim char As String
    Dim codepoint As Long
    Dim charConverted As String
    Dim i As Long
    Dim state As UnicodeToVBACodeConverterState
    Dim currLineLength As Long
    
    currLineLength = 0
    state = Start
    
    If Len(text) = 0 Then Exit Function
    
    For i = 1 To Len(text)
        char = Mid$(text, i, 1)
        codepoint = AscWLong(char)
        Select Case codepoint
            Case &H22
                char = """""" ' works out to 2 double-quote characters
                Select Case state
                    Case Start
                        result = """" & char
                        currLineLength = Len(result)
                    Case InQuote
                        If currLineLength + 5 >= VBA_MAX_LINE_LENGTH Then
                            result = result & """ _" & vbCrLf & "    """ & char
                            currLineLength = 7
                        Else
                            result = result & char
                            currLineLength = currLineLength + 2
                        End If
                    Case NotInQuote
                        If currLineLength + 9 >= VBA_MAX_LINE_LENGTH Then
                            currLineLength = 7
                        Else
                            result = result & " & """ & char
                            currLineLength = currLineLength + 6
                        End If
                End Select
                state = InQuote
            Case &H20, &H21, &H23 To &H7E
                Select Case state
                    Case Start
                        result = """" & char
                        currLineLength = Len(result)
                    Case InQuote
                        If currLineLength + 4 >= VBA_MAX_LINE_LENGTH Then
                            result = result & """ _" & vbCrLf & "    """ & char
                            currLineLength = 6
                        Else
                            result = result & char
                            currLineLength = currLineLength + 1
                        End If
                    Case NotInQuote
                        If currLineLength + 8 >= VBA_MAX_LINE_LENGTH Then
                            currLineLength = 6
                        Else
                            result = result & " & """ & char
                            currLineLength = currLineLength + 5
                        End If
                End Select
                state = InQuote
            Case Else
                Select Case state
                    Case Start
                        charConverted = CodepointToVBACode(codepoint)
                    Case InQuote
                        result = result & """"
                        charConverted = " & " & CodepointToVBACode(codepoint)
                    Case NotInQuote
                        charConverted = " & " & CodepointToVBACode(codepoint)
                End Select
                If currLineLength + Len(charConverted) + 4 >= VBA_MAX_LINE_LENGTH Then
                    result = result & " _" & vbCrLf & "   "
                    currLineLength = 0
                End If
                result = result & charConverted
                currLineLength = currLineLength + Len(charConverted)
                state = NotInQuote
        End Select
    Next i
    
    If state = InQuote Then result = result & """"
    ConvertUnicodeTextToVBACode = result
 
End Function

'@Description "Converts a single input codepoint to a string of corresponding VBA code. The resulting string is either a VBA constant name or a call to ChrW$() with a hex codepoint inside."
Public Function CodepointToVBACode(codepoint As Long) As String
Attribute CodepointToVBACode.VB_Description = "Converts a single input codepoint to a string of corresponding VBA code. The resulting string is either a VBA constant name or a call to ChrW$() with a hex codepoint inside."
    Select Case codepoint
        Case 0
            CodepointToVBACode = "vbNullChar"
        Case 8
            CodepointToVBACode = "vbBack"
        Case 9
            CodepointToVBACode = "vbTab"
        Case &HA
            CodepointToVBACode = "vbLf"
        Case &HC
            CodepointToVBACode = "vbFormFeed"
        Case &HD
            CodepointToVBACode = "vbCr"
        Case &HB
            CodepointToVBACode = "vbVerticalTab"
        Case Else
            CodepointToVBACode = "ChrW$(&H" & Hex$(codepoint) & ")"
    End Select
End Function

Public Sub Demo_ConvertUnicodeTextToVBACode()
    With ActiveWindow.Selection.TextRange2
        .Font.Size = 8
        .Font.Name = "Cascadia Code"
        .text = ConvertUnicodeTextToVBACode(.text)
    End With
End Sub

'@Description "Wrapper around AscW to return a Long value that is always positive."
Public Function AscWLong(char As String) As Long
Attribute AscWLong.VB_Description = "Wrapper around AscW to return a Long value that is always positive."
    AscWLong = AscW(char) And &HFFFF&
End Function

'@EntryPoint
'@Description "Opens a prompt for you to type Unicode text that will be converted to valid VBA and inserted at the cursor position in the VB Editor. Note that if you have text selected in the VB Editor, it will be overwritten."
Public Sub PromptToInsertUnicodeStringIntoVBA()
Attribute PromptToInsertUnicodeStringIntoVBA.VB_Description = "Opens a prompt for you to type Unicode text that will be converted to valid VBA and inserted at the cursor position in the VB Editor. Note that if you have text selected in the VB Editor, it will be overwritten."
    Load UnicodeToVBAPrompt
    UnicodeToVBAPrompt.Show vbModeless
End Sub

Private Sub Demo_PromptToInsertUnicodeStringIntoVBA()
    Debug.Print "VBA" & ChrW$(&H306F) & ChrW$(&H672C) & ChrW$(&H5F53) & ChrW$(&H306B) & ChrW$(&H5384) & ChrW$(&H4ECB) & ChrW$(&H306A) & ChrW$(&H30D7) & ChrW$(&H30ED) & ChrW$(&H30B0) & ChrW$(&H30E9) & ChrW$(&H30DF) & ChrW$(&H30F3) & ChrW$(&H30B0) & ChrW$(&H74B0) & ChrW$(&H5883) & ChrW$(&H3002) & ChrW$(&H30A8) & ChrW$(&H30E9) & ChrW$(&H30FC) & ChrW$(&H30CF) & ChrW$(&H30F3) & ChrW$(&H30C9) & ChrW$(&H30EA) & ChrW$(&H30F3) & ChrW$(&H30B0) & ChrW$(&H306E) & ChrW$(&H30B7) & ChrW$(&H30F3) & ChrW$(&H30BF) & ChrW$(&H30C3) _
    & ChrW$(&H30AF) & ChrW$(&H30B9) & ChrW$(&H304C) & ChrW$(&H6C5A) & ChrW$(&H3044) & ChrW$(&H3002) & ChrW$(&H30E6) & ChrW$(&H30CB) & ChrW$(&H30B3) & ChrW$(&H30FC) & ChrW$(&H30C9) & ChrW$(&H306B) & ChrW$(&H306F) & ChrW$(&H5BFE) & ChrW$(&H5FDC) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H3044) & ChrW$(&H308B) & ChrW$(&H3051) & ChrW$(&H3069) & ChrW$(&H3001) & "VBA" & ChrW$(&H30B3) & ChrW$(&H30FC) & ChrW$(&H30C9) & ChrW$(&H306E) & ChrW$(&H4E2D) & ChrW$(&H306B) & ChrW$(&H306F) & "ASCII" & ChrW$(&H5916) _
    & ChrW$(&H306E) & ChrW$(&H6587) & ChrW$(&H5B57) & ChrW$(&H304C) & ChrW$(&H304B) & ChrW$(&H3051) & ChrW$(&H306A) & ChrW$(&H3044) & ChrW$(&H3002) & vbCr & vbLf & ChrW$(&H672C) & ChrW$(&H5F53) & ChrW$(&H306B) & ChrW$(&H9762) & ChrW$(&H5012) & ChrW$(&H304F) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Sub
