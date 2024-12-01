VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZclsTestMultipleStatements"
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

Public Sub TestCustomerList()
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (1, 'a');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (2, 'c');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (3, 'b');"

    Dim cls As clsCustomerList
    Set cls = New clsCustomerList

    Dim Expected As Variant
    Expected = "(Tbl_Customers.Name) Like '*a*' OR (Tbl_Customers.Name) Like '*c*' OR (Tbl_Customers.Name) Like '*b*' "
    Dim Actual As Variant

    Dim Rst As DAO.Recordset
    Actual = cls.CustomerList(Rst)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestSQLText()
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (1, 'a');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (2, 'c');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (3, 'b');"


    Dim Expected As Variant
    Expected = "SELECT Tbl_Customers.Name , Tbl_Invoices.Date , Tbl_Invoices.Document_Number, Tbl_Invoices.Amount , Tbl_Invoices.Overdue_Date , Tbl_Invoices.mEMO FROM Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID WHERE ( ( (Tbl_Customers.Name) Like '*a*' OR (Tbl_Customers.Name) Like '*c*' OR (Tbl_Customers.Name) Like '*b*' ) AND ( ( Tbl_Invoices.Update_date)=Date())) ORDER BY Tbl_Customers.Name, Tbl_Invoices.Date;"
    Dim Actual As Variant

    Dim Rst As DAO.Recordset
    'Set Rst = CurrentProject.Connection.Execute("select * from Tbl_CustomersList")
    Dim cls As clsCustomerList
    Set cls = New clsCustomerList
    Actual = cls.SQLText(Rst)
    Dim temp As String
    Do
        temp = Actual
        Actual = Replace(Actual, Space(2), Space(1))
    Loop Until temp = Actual

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub
Public Sub TestMakeExcelFileNoInvoices()
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (1, 'a');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (2, 'c');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (3, 'b');"

    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'a', 'a', #28/11/2024 22:27:27#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'b', 'b', #28/11/2024 22:27:29#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'c', 'c', #28/11/2024 22:27:30#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'd', 'd', #28/11/2024 22:27:30#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"


    'CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment]) VALUES (661554, 'a', #28/11/2024#, True, False);"
    'CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment]) VALUES (661555, 'b', #28/11/2024#, True, False);"
    'CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment]) VALUES (661556, 'c', #28/11/2024#, True, False);"

    Dim Expected As Variant
    'Expected = "SELECT Tbl_Customers.Name , Tbl_Invoices.Date , Tbl_Invoices.Document_Number, Tbl_Invoices.Amount , Tbl_Invoices.Overdue_Date , Tbl_Invoices.mEMO FROM Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID WHERE ( ( (Tbl_Customers.Name) Like '*a*' OR (Tbl_Customers.Name) Like '*c*' OR (Tbl_Customers.Name) Like '*b*' ) AND ( ( Tbl_Invoices.Update_date)=Date())) ORDER BY Tbl_Customers.Name, Tbl_Invoices.Date;"
    Dim Actual As Variant

    Dim Rst As DAO.Recordset
    'Set Rst = CurrentProject.Connection.Execute("select * from Tbl_CustomersList")
    Dim cls As clsCustomerList
    Set cls = New clsCustomerList
    Actual = cls.SQLText(Rst)
    Dim temp As String
    Do
        temp = Actual
        Actual = Replace(Actual, Space(2), Space(1))
    Loop Until temp = Actual

    Dim Path As String
    Path = Application.CurrentProject.Path & "\test.xlsx"
    'Path = "C:\Users\alberto\OneDrive\Credit Hunter\test.xlsx"
    On Error Resume Next
    Kill Path
    On Error GoTo 0
    Rst.MoveFirst
    Assert.IsFalse cls.MakeExcelFile(Rst, Path)
    'Assert.That Actual, Iz.EqualTo(Expected)
End Sub



Public Sub TestMakeExcelFileWithInvoices()
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (1, 'a');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (2, 'c');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (3, 'b');"

    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'a', 'a', #28/11/2024 22:27:27#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'b', 'b', #28/11/2024 22:27:29#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'c', 'c', #28/11/2024 22:27:30#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"
    CurrentDb.Execute "INSERT INTO Tbl_Customers(Customer_code, Name, NextAppointment, ToSendStatement, ToSendRequestRelease, StatementForm, RetailOEM, ToReleaseOrder, Language, FacturaNumberToBePrinted, PullTicketNumberToBePrinted, OriginalInvoiceAmountToBePrinted, [Credit Limit]) VALUES (  'd', 'd', #28/11/2024 22:27:30#, False, False, 0, 'Retail', False, 1, False, False, False, 0);"

    Dim aDate As String
    aDate = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
    CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment],[amount]) VALUES (661554, 'a', #" & aDate & "#, True, False,1);"
    CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment],[amount]) VALUES (661555, 'b', #" & aDate & "#, True, False,2);"
    CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment],[amount]) VALUES (661556, 'c', #" & aDate & "#, True, False,3);"
    CurrentDb.Execute "INSERT INTO Tbl_Invoices([ID], [Customer_ID], [Update_date], [QueryToBePrinted], [Attachment],[amount]) VALUES (661556, 'd', #" & aDate & "#, True, False,4);"

    Dim Expected As Variant
    Dim Actual As Variant

    Dim cls As clsCustomerList
    Set cls = New clsCustomerList
    Dim Rst As DAO.Recordset
    Call cls.GetRst(Rst)
    Actual = cls.SQLText(Rst)
    Dim temp As String
    Do
        temp = Actual
        Actual = Replace(Actual, Space(2), Space(1))
    Loop Until temp = Actual

    Dim Path As String
    Path = Application.CurrentProject.Path & "\tests\test.xlsx"

    On Error Resume Next
    Kill Path

    On Error GoTo 0
    Rst.MoveFirst
    Assert.IsTrue cls.MakeExcelFile(Rst, Path)
End Sub
