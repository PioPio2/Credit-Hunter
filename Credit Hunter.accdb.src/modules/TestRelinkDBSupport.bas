Attribute VB_Name = "TestRelinkDBSupport"
''@TestModule
''@Folder("Tests")
'
'Option Compare Database
'
'Option Explicit
'Option Private Module
'
'Private Assert As Rubberduck.AssertClass
'Private Fakes As Rubberduck.FakesProvider
'
''@ModuleInitialize
'Private Sub ModuleInitialize()
'    'this method runs once per module.
'    Set Assert = New Rubberduck.AssertClass
'    Set Fakes = New Rubberduck.FakesProvider
'End Sub
'
''@ModuleCleanup
'Private Sub ModuleCleanup()
'    'this method runs once per module.
'    Set Assert = Nothing
'    Set Fakes = Nothing
'End Sub
'
''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub
'
''@TestMethod("Uncategorized")
'Private Sub TestMethod1()                        'TODO Rename test
'    On Error GoTo TestFail
'
'    Dim clsLink   As clsRelinkDB
'    Set clsLink = New clsRelinkDB
'    Dim Result As String
'    Result = clsLink.ReLinkDB
'    Assert.IsTrue Result
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & err.number & " - " & err.Description
'    Resume TestExit
'End Sub
