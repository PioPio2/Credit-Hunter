dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Invoices.Currency, Sum(Tbl_Invoices.Amou"
    "nt) AS ARExposure, Sum(IIf([Overdue_Date]<Date(),[amount],0)) AS AROverdue, Tbl_"
    "CL.CreditLimit AS [CreditLimit in EUR]\015\012FROM Tbl_Customers LEFT JOIN (Tbl_"
    "Invoices LEFT JOIN Tbl_CL ON Tbl_Invoices.Customer_ID = Tbl_CL.Customer_code) ON"
    " Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID\015\012WHERE (((Tbl_Invo"
    "ices.Update_date)=Date()))\015\012GROUP BY Tbl_Customers.Customer_code, Tbl_Invo"
    "ices.Currency, Tbl_CL.CreditLimit;\015\012"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ARExposure"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AROverdue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CreditLimit in EUR"
        dbLong "AggregateType" ="-1"
    End
End
