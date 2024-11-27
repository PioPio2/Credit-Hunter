Attribute VB_Name = "TESTOutlook"
''@TestModule
''@Folder("Tests")
'
'Option Compare Database
'
'Option Explicit
'Option Private Module
'
''Early Binding
'Private Assert As Rubberduck.PermissiveAssertClass
'
'Private Fakes As Object
'
''@ModuleInitialize
'Private Sub ModuleInitialize()
'    'this method runs once per module.
'    Set Assert = CreateObject("Rubberduck.AssertClass")
'    Set Fakes = CreateObject("Rubberduck.FakesProvider")
'
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
'    If GeneralWD Is Nothing Then
'        Set GeneralWD = New Word.Application
'    End If
'End Sub
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'    GeneralWD.Quit
'    Set GeneralWD = Nothing
'End Sub
'
''@TestMethod("Uncategorized")
'Private Sub TestMethod1()                        '
'    On Error GoTo TestFail
'
'    Dim OL As IOutlook
'    Set OL = New clsOutlook
'    OL.CreateOutlook
'
'    Dim Header As clsCustomerHeader
'    Set Header = New clsCustomerHeader
'    Call Header.Populate("Customer Name")
'
'    Dim Details As clsCustomerDetails
'    Set Details = New clsCustomerDetails
'
'    Dim Result As Boolean
'    Dim Template As String
'
'
'    Template = "E:\MS Access\Projects\Credit Hunter\Templates\Word\test.docx"
'    Dim stopwatch As clsTimer
'    Set stopwatch = New clsTimer
'    stopwatch.StartTimer
'
'    Dim a As Collection
'    Set a = Nothing
'    Result = OL.SendEmailFromTemplate("", Header, Details, a, Template, False)
'    stopwatch.StopAndShowTimer
'    'Debug.Print Timer - startTime
'    Assert.AreEqual Result, True
'    Assert.AreEqual OL.NAttachment, 0
'    Assert.AreEqual OL.SendImmediately, False
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
'
''@TestMethod("Uncategorized")
'Private Sub TestMethodWithAttachment()                        '
'    On Error GoTo TestFail
'
'    Dim OL As IOutlook
'    Set OL = New clsOutlook
'    OL.CreateOutlook
'
'    Dim Header As clsCustomerHeader
'    Set Header = New clsCustomerHeader
'    Call Header.Populate("Customer Name")
'
'    Dim Details As clsCustomerDetails
'    Set Details = New clsCustomerDetails
'
'    Dim Result As Boolean
'    Dim Template As String
'
'
'    Template = "E:\MS Access\Projects\Credit Hunter\Templates\Word\test.docx"
'    Dim stopwatch As clsTimer
'    Set stopwatch = New clsTimer
'    stopwatch.StartTimer
'
'    Dim Attach As Collection
'    Set Attach = New Collection
'    Attach.Add "E:\MS Access\Projects\Credit Hunter\Templates\Word\TestAttachment.txt"
'    Result = OL.SendEmailFromTemplate("", Header, Details, Attach, Template, False)
'    stopwatch.StopAndShowTimer
'    'Debug.Print Timer - startTime
'    Assert.AreEqual Result, True
'    Assert.AreEqual OL.NAttachment, 1
'    Assert.AreEqual OL.SendImmediately, False
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
'
'
