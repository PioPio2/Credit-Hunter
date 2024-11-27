dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Invoices.Currency, Tbl_Currencies.Exchan"
    "geRate, Sum(IIf([tbl_invoices].[update_date]=Date(),IIf([tbl_invoices].[overdue_"
    "date]<=Date()-15,[Amount],0),0))*[ExchangeRate] AS Expr1\015\012FROM (Tbl_Custom"
    "ers LEFT JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Custome"
    "r_ID) LEFT JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.Currenc"
    "yID\015\012GROUP BY Tbl_Customers.Customer_code, Tbl_Invoices.Currency, Tbl_Curr"
    "encies.ExchangeRate;\015\012"
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
        dbText "Name" ="Expr1"
        dbText "Format" ="Standard"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
End
