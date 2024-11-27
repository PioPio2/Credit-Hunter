Attribute VB_Name = "TestCustomerHeaderSupportLIVE"
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
'Private Fakes As Object
'
''@ModuleInitialize
'Private Sub ModuleInitialize()
'    'this method runs once per module.
'    Set Assert = CreateObject("Rubberduck.AssertClass")
'    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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
'    Dim CustomerID As String
'    CustomerID = "123"
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Customers")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Customers ( Customer_code, Credit_controller, Name, Address, Address2, Address3, Address4, Address5, Country, Update_date, OWN_company, OWN_bank_details1, OWN_bank_details2, OWN_bank_details3, OWN_bank_details4, NextAppointment, Email, [DA TOGLIEREEEEEEEE TextEmail], ccEmail, StatusDate, ToSendStatement, [Index], [Note], ToSendRequestRelease, TotalInsurance, StatementForm, RetailOEM, ToReleaseOrder, Status, Timezone, [Language], ContactNames, DSO, EmailCode, Area, LastStatementSent, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, HighestExposure, ReleaseNotes, MonthlyTargetInMainCurrency, " & _
'    "[Credit Limit], MainPhoneNumber )SELECT 'Strike Electrical Distributors Limited' AS Expr1, 1 AS Expr2, 'Strike Electrical Distributors Limited' AS Expr3, '245 Green Lane Walsall West Midlands ' AS Expr4, 0 AS Expr5, 0 AS Expr6, 0 AS Expr7, 0 AS Expr8, 0 AS Expr9, 0 AS Expr10, 0 AS Expr11, 0 AS Expr12, 0 AS Expr13, 0 AS Expr14, 0 AS Expr15, '17/07/2024' AS Expr16, 'tony@strikeelectrical.co.uk' AS Expr17, 0 AS Expr18, 0 AS Expr19, 0 AS Expr20, False AS Expr21, 0 AS Expr22, 0 AS Expr23, False AS Expr24, 0 AS Expr25, 0 AS Expr26, 'Retail' AS Expr27, False AS Expr28, 0 AS Expr29, 0 AS Expr30, 1 AS Expr31, 0 AS Expr32, 0 AS Expr33, 0 AS Expr34, 0 AS Expr35, 0 AS Expr36, False AS Expr37, False AS Expr38, False AS Expr39, 0 AS Expr40, 0 AS Expr41, 0 AS Expr42, 50000 AS Expr43, 0 AS Expr44;")
'
'    Dim Header As clsCustomerHeaderSupportLIVE
'    Set Header = New clsCustomerHeaderSupportLIVE
'    Dim Dict As Scripting.Dictionary
'    Set Dict = Header.ICustomerHeaderSupport_Populate(CustomerID)
'    Assert.AreEqual Dict("Name"), "Strike Electrical Distributors Limited"
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & err.number & " - " & err.Description
'    Resume TestExit
'End Sub
