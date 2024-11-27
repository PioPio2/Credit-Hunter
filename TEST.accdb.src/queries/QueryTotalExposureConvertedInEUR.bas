dbMemo "SQL" ="SELECT DISTINCTROW Tbl_Invoices.Customer_ID, Sum(([Amount]*[ExchangeRate])) AS T"
    "otalExposureConvertedInEUR\015\012FROM (Tbl_Customers INNER JOIN Tbl_Invoices ON"
    " Tbl_Customers.Customer_code=Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies"
    " ON Tbl_Invoices.Currency=Tbl_Currencies.CurrencyID\015\012WHERE (((Tbl_Invoices"
    ".Update_date)=Date()))\015\012GROUP BY Tbl_Invoices.Customer_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="TotalExposureConvertedInEUR"
        dbLong "AggregateType" ="-1"
    End
End
