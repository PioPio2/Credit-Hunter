dbMemo "SQL" ="SELECT Tbl_Invoices.Customer_ID, Sum([Amount]/[ExchangeRateToMainCurrency]) AS T"
    "otalOverdueOver90\015\012FROM Tbl_Invoices LEFT JOIN QueryCurrentExchangeRatesTo"
    "MainCurrency ON Tbl_Invoices.Currency = QueryCurrentExchangeRatesToMainCurrency."
    "OriginalCurrency\015\012WHERE (((Tbl_Invoices.Update_date)=Date()) AND ((Tbl_Inv"
    "oices.Overdue_Date)<=DateAdd('d',-90,[MonthEnd])))\015\012GROUP BY Tbl_Invoices."
    "Customer_ID\015\012ORDER BY Tbl_Invoices.Customer_ID;\015\012"
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
        dbText "Name" ="Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalOverdueOver90"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
