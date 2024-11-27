dbMemo "SQL" ="SELECT Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashColl"
    "ected.[Payment Date], Sum(Tbl_CashCollected.Amount) AS [Amount in EUR], Tbl_Cash"
    "Collected.Currency, Tbl_CashCollected.[Original amount], Tbl_Users.ID\015\012FRO"
    "M ((Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = "
    "Tbl_CashCollected.CustomerID) INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Cur"
    "rency = Tbl_Currencies.CurrencyID) INNER JOIN Tbl_Users ON Tbl_Customers.Credit_"
    "controller = Tbl_Users.ID\015\012GROUP BY Tbl_Users.Name, Tbl_Customers.RetailOE"
    "M, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.Curre"
    "ncy, Tbl_CashCollected.[Original amount], Tbl_Users.ID\015\012HAVING (((Tbl_Cash"
    "Collected.[Payment Date])>=#02/21/15# And (Tbl_CashCollected.[Payment Date])<=#0"
    "3/31/15#))\015\012ORDER BY Sum(Tbl_CashCollected.Amount) DESC;\015\012"
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
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.[Payment Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Amount in EUR"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.RetailOEM"
        dbLong "AggregateType" ="-1"
    End
End
