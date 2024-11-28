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
Public Function SQLText(CustomerSelectRst As Variant) As String
'    Dim customers  As String
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
    '    SQL = SQL & "                        ( "
    '    SQL = SQL & "                                Tbl_Customers.Name) Like ' * abc * ' "
    '    SQL = SQL & "                        Or ( "
    '    SQL = SQL & "                                Tbl_Customers.Name) Like ' * Def * ' ) "
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

Public Function CustomerList(rst As Variant) As String
    'Dim Rst           As ADODB.Recordset
    Dim vDat          As Variant

    'Set Rst = CurrentProject.Connection.Execute("select * from tblTemp4")
    vDat = rst.GetRows
    CustomerList = ""
    If UBound(vDat, 2) > 0 Then
        Dim I As Long
        For I = 0 To UBound(vDat, 2)
            CustomerList = CustomerList & "(Tbl_Customers.Name) Like '*" & vDat(1, I) & "*' OR "
        Next I
        CustomerList = Left(CustomerList, Len(CustomerList) - 3)
    End If

End Function
Public Function MakeExcelFile(CustomerSelectRst As Variant, Path As String) As Boolean
    Const DummyQuery As String = "QueryMultipleStatements"
    If Not CustomerSelectRst.EOF Then
        Dim qd As QueryDef
        Set qd = CurrentDb.QueryDefs(DummyQuery)
        qd.SQL = SQLText(CustomerSelectRst)
        qd.OpenRecordset
        If DCount("*", DummyQuery) > 0 Then
            On Error GoTo Err
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, DummyQuery, Path, True
            MakeExcelFile = True
        Else
            MakeExcelFile = False
            Exit Function
        End If
        On Error GoTo 0
        MakeExcelFile = True
Else
    MakeExcelFile = False
End If
Exit Function

Err:
    MakeExcelFile = False
    On Error GoTo 0
End Function
