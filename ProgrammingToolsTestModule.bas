Attribute VB_Name = "ProgrammingToolsTestModule"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("Uncategorized")
Private Sub TestCodepointToVBACode()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(0), "vbNullChar"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(8), "vbBack"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(9), "vbTab"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(&HA), "vbLf"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(&HC), "vbFormFeed"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(&HD), "vbCr"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(&HB), "vbVerticalTab"
    Assert.AreEqual ProgrammingTools.CodepointToVBACode(&H3000), "ChrW$(&H3000)"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestConvertUnicodeTextToVBACode()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim inputs As New Collection
    Dim outputs As New Collection
    Dim expectedOutputs As New Collection
    Dim sld As Slide
    Dim shp As Shape
    Dim testTable As Table
    Dim i As Long
    
    'get the table containing the relevant results
    Set sld = ActivePresentation.Slides.Item(8)
    For Each shp In sld.Shapes.Range
        If shp.HasTable Then Set testTable = shp.Table
    Next shp
    
    For i = 2 To testTable.Rows.Count
        inputs.Add testTable.Rows(i).Cells(2).Shape.TextFrame2.TextRange.text
        expectedOutputs.Add testTable.Rows(i).Cells(3).Shape.TextFrame2.TextRange.text
    Next i
    
    'Act:
    For i = 1 To inputs.Count
        outputs.Add ProgrammingTools.ConvertUnicodeTextToVBACode(inputs.Item(i))
    Next i

    'Assert:
    Assert.IsNotNothing testTable, "Failed to find test data"
    For i = 1 To inputs.Count
        Assert.AreEqual outputs.Item(i), expectedOutputs.Item(i), "Failed on input: " & inputs.Item(i) & " | Expected: " & expectedOutputs.Item(i) & "| Output: " & outputs.Item(i)
    Next i
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAscWLong()
    On Error GoTo TestFail
    
    'Arrange:
    Dim inputs As New Collection
    Dim outputs As New Collection
    Dim expectedOutputs As New Collection
    Dim sld As Slide
    Dim shp As Shape
    Dim testTable As Table
    Dim i As Long
    
    'get the table containing the relevant results
    Set sld = ActivePresentation.Slides.Item(9)
    For Each shp In sld.Shapes.Range
        If shp.HasTable Then Set testTable = shp.Table
    Next shp
    
    inputs.Add vbNullChar
    expectedOutputs.Add 0&
    
    inputs.Add "A"
    expectedOutputs.Add &H41&
    
    inputs.Add testTable.Rows(4).Cells(2).Shape.TextFrame2.TextRange.text
    expectedOutputs.Add &H3001&
    
    inputs.Add testTable.Rows(5).Cells(2).Shape.TextFrame2.TextRange.text
    expectedOutputs.Add &HFFE5&
    
    'Act:
    For i = 1 To inputs.Count
        outputs.Add ProgrammingTools.AscWLong(inputs.Item(i))
    Next i

    'Assert:
    Assert.IsNotNothing testTable, "Failed to find test data"
    For i = 1 To inputs.Count
        Assert.AreEqual outputs.Item(i), expectedOutputs.Item(i), "Failed on input: " & inputs.Item(i)
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub TestConvertChrWCallsToUnicode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim inputs As New Collection
    Dim outputs As New Collection
    Dim expectedOutputs As New Collection
    Dim sld As Slide
    Dim shp As Shape
    Dim testTable As Table
    Dim i As Long
    
    'get the table containing the relevant results
    Set sld = ActivePresentation.Slides.Item(10)
    For Each shp In sld.Shapes.Range
        If shp.HasTable Then Set testTable = shp.Table
    Next shp
    
    For i = 2 To testTable.Rows.Count
        inputs.Add testTable.Rows(i).Cells(2).Shape.TextFrame2.TextRange.text
        expectedOutputs.Add testTable.Rows(i).Cells(3).Shape.TextFrame2.TextRange.text
    Next i
    
    'Act:
    For i = 1 To inputs.Count
        outputs.Add ProgrammingTools.ConvertChrWCallsToUnicode(inputs.Item(i))
    Next i

    'Assert:
    Assert.IsNotNothing testTable, "Failed to find test data"
    For i = 1 To inputs.Count
        Assert.AreEqual outputs.Item(i), expectedOutputs.Item(i), "Failed on input: " & inputs.Item(i) & " | Expected: " & expectedOutputs.Item(i) & "| Output: " & outputs.Item(i)
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
