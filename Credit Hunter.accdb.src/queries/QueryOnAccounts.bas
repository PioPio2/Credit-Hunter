dbMemo "SQL" ="SELECT Tbl_Customers.Credit_controller, Tbl_Customers.Customer_code, Tbl_Custome"
    "rs.Name, Tbl_Invoices.Document_Number, Tbl_Invoices.Date, Tbl_Invoices.Currency,"
    " Sum(Tbl_Invoices.Amount) AS SumOfAmount\015\012FROM Tbl_Customers INNER JOIN Tb"
    "l_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID\015\012WHER"
    "E (((Tbl_Invoices.Type)=3) AND ((Tbl_Invoices.Update_date)=(SELECT Max(Tbl_Invoi"
    "ces.Update_date) AS MaxOfUpdate_date\015\012FROM Tbl_Invoices;)))\015\012GROUP B"
    "Y Tbl_Customers.Credit_controller, Tbl_Customers.Customer_code, Tbl_Customers.Na"
    "me, Tbl_Invoices.Document_Number, Tbl_Invoices.Date, Tbl_Invoices.Currency\015\012"
    "HAVING (((Tbl_Customers.Credit_controller)=GetNumCreditController(fOSUserName())"
    "))\015\012ORDER BY Sum(Tbl_Invoices.Amount);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Document_Number"
        dbLong "AggregateType" ="-1"
    End
End
