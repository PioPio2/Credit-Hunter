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

   Dim rst           As ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("select * from Tbl_CustomersList")
   Actual = cls.CustomerList(rst)

   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestSQLText()
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (1, 'a');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (2, 'c');"
    CurrentDb.Execute "INSERT INTO Tbl_CustomersList(ID, Customer) VALUES (3, 'b');"


   Dim Expected As Variant
   'Expected = "(Tbl_Customers.Name) Like '*a*' OR (Tbl_Customers.Name) Like '*c*' OR (Tbl_Customers.Name) Like '*b*' "
   Dim Actual As Variant

   Dim rst           As ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("select * from Tbl_CustomersList")
    rst.MoveLast
    rst.MoveFirst
   'Actual = cls.CustomerList(rst)
   Dim cls As clsCustomerList
   Set cls = New clsCustomerList
   Actual = cls.SQLText(rst)

   Assert.That Actual, Iz.EqualTo(Expected)
End Sub
