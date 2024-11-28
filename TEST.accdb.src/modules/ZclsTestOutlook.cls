VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZclsTestOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'AccUnit:TestClass

Public Sub Setup()
' This procedure is called before each test
    Call DeleteTables
End Sub

Public Sub Teardown()
' This procedure is called after each test
End Sub

Public Sub CreateOutlook()

'Call IOutlook_CreateOutlook
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, [E-mailAddress]) VALUES (1, 'alberto', 'Alberto', 'alberto.paganini70@gmail.com');"
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, Querywithoutcreditcontroller, Onaccountsstillopen, Whopaidyesterdayroutine, Superuser, EmailSentToSender) VALUES (2, 'aaa', 'admin', False, False, False, False, False);"

   Dim App As clsOutlook
   Set App = New clsOutlook
   Dim Expected As Boolean
   Expected = True
   Dim Actual As Boolean
   Actual = App.IOutlook_CreateOutlook
   'Const Expected As Long = vbMonday

   'Actual = Weekday(#1/1/2024#, vbSunday)

   Assert.That Actual, Iz.EqualTo(Expected)
   Assert.IsNotNull OL
End Sub

Public Sub SendEmailNoAttach()
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, [E-mailAddress]) VALUES (1, 'alberto', 'Alberto', 'alberto.paganini70@gmail.com');"
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, Querywithoutcreditcontroller, Onaccountsstillopen, Whopaidyesterdayroutine, Superuser, EmailSentToSender) VALUES (2, 'aaa', 'admin', False, False, False, False, False);"

    Dim App As clsOutlook
    Set App = New clsOutlook
    Dim Expected As Boolean
    Expected = True
    Dim Actual As Boolean

    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Credit_controller, Name, Address, NextAppointment, Email, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES ('Strike Electrical Distributors Limited', 1, 'Strike Electrical Distributors Limited', '245 Green Lane Walsall West Midlands', #17/07/2024#, 'mytest@gmmail.com', False, False, 0, 'Retail', False, 1, False, False, False, 50000);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code) VALUES ('abc');"


    Dim Header As clsCustomerHeader
    Set Header = New clsCustomerHeader
    Call Header.Populate("Strike Electrical Distributors Limited")

    Dim Details As clsCustomerDetails
    Set Details = New clsCustomerDetails
    Call Details.Populate(2, 1, "USD")


    Dim Template As String
    Template = Application.CurrentProject.Path()
    Template = Template & "\Templates\Word\test.docx"
    Dim stopwatch As clsTimer
    Set stopwatch = New clsTimer
    stopwatch.StartTimer

    Dim a As Collection
    Set a = Nothing
    Dim Result As Boolean
    Result = App.IOutlook_SendEmailFromTemplate("Test Subject", Header, Details, a, Template, False)
    stopwatch.StopAndShowTimer

    Assert.That Result, Iz.EqualTo(True)
    Assert.AreEqual App.IOutlook_NAttachment, 0
    Assert.IsNotNull OL
End Sub

Public Sub SendEmailWithAttach()
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, [E-mailAddress]) VALUES (1, 'alberto', 'Alberto', 'alberto.paganini70@gmail.com');"
    CurrentDb.Execute "INSERT INTO Tbl_Users(ID, UserName, Name, Querywithoutcreditcontroller, Onaccountsstillopen, Whopaidyesterdayroutine, Superuser, EmailSentToSender) VALUES (2, 'aaa', 'admin', False, False, False, False, False);"

    Dim App As clsOutlook
    Set App = New clsOutlook
    Dim Expected As Boolean
    Expected = True
    Dim Actual As Boolean

    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Credit_controller, Name, Address, NextAppointment, Email, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES ('Strike Electrical Distributors Limited', 1, 'Strike Electrical Distributors Limited', '245 Green Lane Walsall West Midlands', #17/07/2024#, 'mytest@gmmail.com', False, False, 0, 'Retail', False, 1, False, False, False, 50000);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code) VALUES ('abc');"

    Dim Header As clsCustomerHeader
    Set Header = New clsCustomerHeader
    Call Header.Populate("Strike Electrical Distributors Limited")
'
    Dim Details As clsCustomerDetails
    Set Details = New clsCustomerDetails
    Call Details.Populate(2, 1, "USD")


    Dim Template As String
    Template = Application.CurrentProject.Path()
    Template = Template & "\Templates\Word\test.docx"
    Dim stopwatch As clsTimer
    Set stopwatch = New clsTimer
    stopwatch.StartTimer

    Dim a As Collection
    Set a = New Collection
    a.Add "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Test attachment.txt"
    Dim Result As Boolean
    Result = App.IOutlook_SendEmailFromTemplate("Test Subject", Header, Details, a, Template, False)
    stopwatch.StopAndShowTimer

    Assert.AreEqual App.IOutlook_NAttachment, 1
    Assert.That Result, Iz.EqualTo(True)
    Assert.IsNotNull OL
End Sub




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
'

'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'    If GeneralWD Is Nothing Then
'        Set GeneralWD = New Word.Application
'    End If
'End Sub
'

'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'    GeneralWD.Quit
'    Set GeneralWD = Nothing
'End Sub
'

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
'    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
'    Resume TestExit
'End Sub
'

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
'    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
'    Resume TestExit
'End Sub
'
'
'
