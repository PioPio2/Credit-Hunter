dbMemo "SQL" ="SELECT DISTINCTROW Tbl_Invoices_History.Customer_ID, Sum([Amount]*[ExchangeRate]"
    ") AS TotalInvoicedConvertedInEUR\015\012FROM Tbl_Currencies INNER JOIN (Tbl_Cust"
    "omers INNER JOIN Tbl_Invoices_History ON Tbl_Customers.Customer_code=Tbl_Invoice"
    "s_History.Customer_ID) ON Tbl_Currencies.CurrencyID=Tbl_Invoices_History.Currenc"
    "y\015\012GROUP BY Tbl_Invoices_History.Customer_ID;\015\012"
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
        dbText "Name" ="TotalInvoicedConvertedInEUR"
        dbLong "AggregateType" ="-1"
    End
End
