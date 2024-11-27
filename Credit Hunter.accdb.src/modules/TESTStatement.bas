Attribute VB_Name = "TESTStatement"
''@TestModule
''@Folder("Tests")
'
'Option Compare Database
'
'Option Explicit
'Option Private Module
'
''Private Assert As Rubberduck.AssertClass
''Early Binding
'Private Assert As Rubberduck.PermissiveAssertClass
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
'Private Sub TestStatementBasic()
'    On Error GoTo TestFail
'
'    Dim StatementSupport As iStatementSupport
'    Set StatementSupport = New clsStatementSupportLIVE
'
'    Dim CustomerHeader As clsCustomerHeader
'    Set CustomerHeader = New clsCustomerHeader
'    Call CustomerHeader.Populate("customerid")
'
''    Dim HeaderTags As Collection
' '   Set HeaderTags = New Collection
'  '  HeaderTags.Add "<<Customer Name>>", CustomerHeader.getCustomerName
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'ABM Electrical Wholesale Ltd', '16/07/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'ABM Electrical Wholesale Ltd', '16/07/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'ABM Electrical Wholesale Ltd', '20/08/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'ABM Electrical Wholesale Ltd', '20/08/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim Statement As IStatement
'    Set Statement = New clsStatement
'    Dim Result As Boolean
'    Dim Overdue As Currency
'    Dim Outstanding As Currency
'    Overdue = 0
'    Outstanding = 0
'    Result = Statement.CreateStatement(StatementSupport, CustomerHeader, Outstanding, Overdue, False)
'    Assert.IsTrue Result                         ' check the statement is created
'
'    Dim Rng As Range
'    Set Rng = GeneralExcel.Sheets(1).Cells.Find(CustomerHeader.getCustomerName)
'    Assert.AreEqual Rng.value, CustomerHeader.getCustomerName ' check the customer name in the statement corresponds to CustomerHeader.getCustomerName passed to the sub
'
'    Dim NRows As Integer
'    NRows = GeneralExcel.Sheets(1).Range("a16").CurrentRegion.Rows.Count
'    Assert.AreEqual NRows, CInt(5)               ' check the n# of lines in the statement is 5 as per db population above (4 data + 1 heading)
'
'
'    GeneralExcel.DisplayAlerts = False
'    GeneralExcel.Workbooks.Close
'    GeneralExcel.DisplayAlerts = True
'    GeneralExcel.Quit
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    GeneralExcel.DisplayAlerts = False
'    GeneralExcel.Workbooks.Close
'    GeneralExcel.DisplayAlerts = True
'    GeneralExcel.Quit
'
'    Assert.Fail "Test raised an error: #" & err.number & " - " & err.Description
'    Resume TestExit
'End Sub
'
