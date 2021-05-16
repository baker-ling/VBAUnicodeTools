# VBAUnicodeTools
 Tools for working with Unicode in the VB Editor

This repository contains a standard VBA module ProgrammingTools.bas and associated userforms for working with Unicode text in the VB Editor (VBE).

## Main Features
There are two main subroutines to help you edit VBA code that deals with Unicode characters outside the ASCII range to circumvent VBE's lack of support for Unicode.

1. `Public Sub PromptToInsertUnicodeStringIntoVBA()`
2. `Public Sub DisplayUnicodeFromVBACode()`

`Public Sub PromptToInsertUnicodeStringIntoVBA()`
Place the cursor wherever you want to insert VBA code for a string containing Unicode characters outside the ASCII range, then run `PromptToInsertUnicodeStringIntoVBA`.
The sub will show a userform with a single textbox that you can type into. Press the "Convert & Insert" button to convert the contents of that textbox to corresponding VBA code and insert it where you placed your cursor in the VB editor.

`Public Sub DisplayUnicodeFromVBACode()`
Select VBA code in the VB editor with Unicode characters hard-coded hexidecimal or decimal codepoints inside `ChrW$` or `ChrW` calls, then run `DisplayUnicodeFromVBACode`. The sub will show a userform with a single textbox containing those `ChrW$`/`ChrW` calls replaced by the Unicode characters they represent.

## Additional Features
There are a few other utility functions for processing Unicode text in VBA.

`Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Microsoft PowerPoint") As VbMsgBoxResult`
A *wide* version of `MsgBox`, adapted from forum post by John_w at https://www.mrexcel.com/board/threads/vba-display-foreign-character-code.1142510/post-5536387.

`Public Sub ShowCharForCodePoint()`
Shows a message box containing the character associated with the codepoint selected in the VB editor. Assumes that only VBA numerical literal is selected.

`Public Function GetSelectedVBACode() As String`
Returns a string containing the code currently selected in the VB editor.

`Public Function ConvertUnicodeTextToVBACode(ByRef text As String) As String`
Takes a string of ordinary text and returns a string with that text converted to corresponding VBA string literals, VBA constants, and calls to calls to `ChrW$`, all concatenated together. VBA line continuations and CRLFs are inserted whenever any line in the output would exceed 512 characters (Adjustable in the private constant `VBA_MAX_LINE_LENGTH` declared at the top of the module. Actual VBE max line length is 1024 characters.).

`Public Function AscWLong(char As String) As Long`
Wrapper around AscW to make sure that all codepoints returned are positive. (Because VBA doesn't have unsigned data types!)

