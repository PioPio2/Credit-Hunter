Attribute VB_Name = "TESTWhoPaidYesterday"
''@TestModule
''@Folder("Tests")
'
'
'Option Compare Database
'
'Option Explicit
'
''Early Binding
'Private Assert As Rubberduck.PermissiveAssertClass
'
'Option Private Module
'
''Private Assert As Object
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
'    Call RelinkTables("E:\MS Access\Projects\Credit Hunter\db\TEST db2.mdb")
'End Sub
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub
'
''@TestMethod("Uncategorized")
'Private Sub TestMethodSetQueryInvoicesBeforeLastDate()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'
'    Call WhoPaidYesterday.Inject(Injection)
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'ABM Electrical Wholesale Ltd', '16/07/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'ABM Electrical Wholesale Ltd', '16/07/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'ABM Electrical Wholesale Ltd', '20/08/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'ABM Electrical Wholesale Ltd', '20/08/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim a As String
'    a = WhoPaidYesterday.BeforeLastInvoiceImportDateSQL
'
'    Dim RS As DAO.Recordset
'    Set RS = Application.CurrentDb.OpenRecordset(a)
'    RS.MoveLast
'    Const ExpectedLong As Long = 2
'    Assert.AreEqual RS.RecordCount, ExpectedLong
'
'    Const Expected As String = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#07/16/2024#));"
'    Assert.AreEqual a, Expected
'
'    Assert.AreEqual CDate(RS.Fields(0).value), CDate(#7/16/2024#)
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
'Private Sub TestMethodUpdateSecondaryQueries()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'
'    Call WhoPaidYesterday.Inject(Injection)
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'ABM Electrical Wholesale Ltd', '16/07/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'ABM Electrical Wholesale Ltd', '16/07/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'ABM Electrical Wholesale Ltd', '20/08/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'ABM Electrical Wholesale Ltd', '20/08/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim Result As Boolean
'    Result = WhoPaidYesterday.UpdateSecondaryQueries
'    Assert.IsTrue Result
'
'    Const Expected As String = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#08/20/2024#));"
'
'    Dim ResultString As String
'    ResultString = Application.CurrentDb.QueryDefs("QueryInvoicesLastDate").SQL
'    '    ResultString = Replace(ResultString, vbLf, "")
'    ResultString = Replace(ResultString, vbCrLf, " ")
'    ResultString = Replace(ResultString, vbCr, "")
'    ResultString = Trim(ResultString)
'    Assert.AreEqual Expected, ResultString
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
'Private Sub TestMethodQueryInvoicesLastDateSQL3()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'
'    Call WhoPaidYesterday.Inject(Injection)
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'Strike Electrical Distributors Limited', '16/07/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'Strike Electrical Distributors Limited', '16/07/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'Strike Electrical Distributors Limited', '20/08/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'Strike Electrical Distributors Limited', '20/08/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim Result As String
'    Const Expected As String = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#08/20/2024#));"
'    Dim RS As DAO.Recordset
'    Result = WhoPaidYesterday.QueryInvoicesLastDateSQL()
'    Assert.AreEqual Result, Expected
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
'Private Sub TestMethodQueryInvoicesLastDateSQL()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'
'    Dim c As Boolean
'    Call WhoPaidYesterday.Inject(Injection)
'
'    Dim Today As Date
'    Dim Yesterday As Date
'    Today = Date
'    Yesterday = Today - 1
'    Dim TodayString As String
'    TodayString = Format(Today, "mm/dd/yyyy")
'    Dim YesterdayString As String
'    YesterdayString = Format(Yesterday, "mm/dd/yyyy")
'
'    '    Call Injection.SetLastInvoiceImportDate("#" & TodayString & "#")
'    '   Call Injection.SetOneBeforeLastInvoiceImportDate("#" & YesterdayString & "#")
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'ABM Electrical Wholesale Ltd', '" & YesterdayString & "', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'ABM Electrical Wholesale Ltd', '" & YesterdayString & "', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'ABM Electrical Wholesale Ltd', '" & TodayString & "', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    ' in this case one invoice is missing, meaning paid --> one customers who paid yesterday
'    'Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'ABM Electrical Wholesale Ltd', '20/08/2024', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim Result As String
'    Const Expected As String = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#09/14/2024#));"
'    Dim RS As DAO.Recordset
'    Result = WhoPaidYesterday.QueryInvoicesLastDateSQL()
'    Assert.AreEqual Result, Expected
'    'Assert.AreEqual RS.Fields(0).value, "ABM Electrical Wholesale Ltd"
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
'Private Sub TestMethodGetWhoPaidYesterday3()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'
'    Dim Today As Date
'    Dim Yesterday As Date
'    Today = Date
'    Yesterday = Today - 1
'    Dim TodayString As String
'    TodayString = Format(Today, "mm/dd/yyyy")
'    Dim YesterdayString As String
'    YesterdayString = Format(Yesterday, "mm/dd/yyyy")
'
'    Call WhoPaidYesterday.Inject(Injection)
'    '    Call Injection.SetLastInvoiceImportDate("#" & TodayString & "#")
'    '   Call Injection.SetOneBeforeLastInvoiceImportDate("#" & YesterdayString & "#")
'
'    Application.CurrentDb.Execute ("DELETE * FROM Tbl_Invoices")
'
'    ' in this case one new invoice is open--> no customers who paid yesterday
'    'Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#])VALUES (661019, 'ABM Electrical Wholesale Ltd', '16/07/2024', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935','' , 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661020, 'ABM Electrical Wholesale Ltd', '" & YesterdayString & "', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661551, 'ABM Electrical Wholesale Ltd', '" & TodayString & "', '21/06/2024', 'INV-Liv30004597', 'ABM Electrical Wholesale Ltd', 'SO-Liv90004935', '', 125.16, 125.16, '21/07/2024', 'GBP',0 ,0 ,0 , True, False,0 ,0 );")
'    Application.CurrentDb.Execute ("INSERT INTO Tbl_Invoices(ID, Customer_ID, Update_date, [Date], Document_Number, Customer_reference, SONumber, Type, OriginalAmount, Amount, Overdue_Date, [Currency], Query, mEMO, QueryDate, QueryToBePrinted, Attachment, CustomsInvoiceNumber, [PullTicketN#]) VALUES (661552, 'ABM Electrical Wholesale Ltd', '" & TodayString & "', '02/06/2023', 'INV-Liv30002254', 'ABM Electrical Wholesale Ltd', 'SO-Liv90002294','' , 22440.67, 2494.66, '02/07/2023', 'GBP', 0, 0,0 , True, False, 0, 0);")
'
'    Dim Result As String
'    Dim Expected As String
'    Expected = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & TodayString & "#));"
'    Dim RS As DAO.Recordset
'    Result = WhoPaidYesterday.QueryInvoicesLastDateSQL()
'    Assert.AreEqual Result, Expected
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
'Private Sub TestMethodGetRecordSource()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Result As String
'    Const Expected As String = "SELECT         DISTINCTROW Tbl_Customers.* FROM         Tbl_Customers INNER JOIN         QueryInvoicesClosedInLastDate ON         Tbl_Customers.Customer_code = QueryInvoicesClosedInLastDate.Customer_ID WHERE         (                 (                         (                                 Tbl_Customers.Credit_controller)=1)); "
'    Result = WhoPaidYesterday.GetRecordSourceSQL
'    Assert.AreEqual Result, Expected
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
'Private Sub TestMethodQueryInvoicesLastDateSQL2()
'    On Error GoTo TestFail
'
'    Dim WhoPaidYesterday As IWhoPaidYesterday
'    Set WhoPaidYesterday = New clsWhoPaidYesterday
'
'    Dim Injection As IWhoPaidYesterdaySupport
'    Set Injection = New clsWhoPaidYesterdaySupportLIVE
'    Call WhoPaidYesterday.Inject(Injection)
'
'    Dim Result As String
'    Const Expected As String = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#09/14/2024#));"
'    Result = WhoPaidYesterday.QueryInvoicesLastDateSQL
'    Assert.AreEqual Result, Expected
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
