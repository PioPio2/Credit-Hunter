VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Sub GetRst(ByRef Rst As DAO.Recordset)
    If Not Rst Is Nothing Then
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
    End If
    Set Rst = CurrentDb.OpenRecordset("select * from Tbl_CustomersList")
End Sub
Public Function SQLText(CustomerSelectRst As DAO.Recordset) As String
    Call GetRst(CustomerSelectRst)
    Dim SQL As String
    SQL = ""
    If Not CustomerSelectRst.EOF > 0 Then
        SQL = SQL & "SELECT "
        SQL = SQL & "        Tbl_Customers.Name          , "
        SQL = SQL & "        Tbl_Invoices.Date           , "
        SQL = SQL & "        Tbl_Invoices.Document_Number, "
        SQL = SQL & "        Tbl_Invoices.Amount         , "
        SQL = SQL & "        Tbl_Invoices.Overdue_Date   , "
        SQL = SQL & "        Tbl_Invoices.mEMO "
        SQL = SQL & "FROM "
        SQL = SQL & "        Tbl_Customers "
        SQL = SQL & "INNER JOIN "
        SQL = SQL & "        Tbl_Invoices "
        SQL = SQL & "ON "
        SQL = SQL & "        Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID "
        SQL = SQL & "WHERE "
        SQL = SQL & "        ( "
        SQL = SQL & "                ( "
        SQL = SQL & CustomerList(CustomerSelectRst)
        SQL = SQL & ")"
        SQL = SQL & "                AND ( "
        SQL = SQL & "                        ( "
        SQL = SQL & "                                Tbl_Invoices.Update_date)=Date())) "
        SQL = SQL & "ORDER BY "
        SQL = SQL & "        Tbl_Customers.Name, "
        SQL = SQL & "        Tbl_Invoices.Date; "
    End If

    SQLText = Trim(SQL)
End Function

Public Function CustomerList(ByRef Rst As DAO.Recordset) As String
    Call GetRst(Rst)
    Rst.MoveLast
    Rst.MoveFirst

    Dim vDat As Variant
    vDat = Rst.GetRows(Rst.RecordCount)

    CustomerList = ""
    If UBound(vDat) > 0 Then
        Dim I As Long
        For I = 0 To UBound(vDat, 2)
            CustomerList = CustomerList & "(Tbl_Customers.Name) Like '*" & vDat(1, I) & "*' OR "
        Next I
        CustomerList = Left(CustomerList, Len(CustomerList) - 3)
    End If
End Function
Public Function MakeExcelFile(CustomerSelectRst As DAO.Recordset, Path As String) As Boolean
    Const DummyQuery As String = "QueryMultipleStatements"
    If Not CustomerSelectRst.EOF Then
        Dim qd As DAO.QueryDef
        Set qd = CurrentDb.QueryDefs(DummyQuery)
        qd.SQL = "select *"
        qd.SQL = SQLText(CustomerSelectRst)
        qd.OpenRecordset

        On Error GoTo Err
        If DCount("*", DummyQuery) > 0 Then
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, DummyQuery, Path, True
            MakeExcelFile = True
        Else
            MakeExcelFile = False
            Exit Function
        End If

        On Error GoTo 0
        '        MakeExcelFile = True
    Else
        MakeExcelFile = False
    End If

    Exit Function

Err:
    MakeExcelFile = False
    On Error GoTo 0
End Function
