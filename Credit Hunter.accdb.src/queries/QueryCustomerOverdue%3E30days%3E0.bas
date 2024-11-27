dbMemo "SQL" ="SELECT Tbl_Users.Name, Tbl_Customers.Name AS [Customer name], Tbl_Customer_Statu"
    "s.Status, Tbl_Customers.StatusDate, DateDiff(\"d\",[StatusDate],Date()) AS [Days"
    " gone by], [QueryCustomerOverdue>30daysOnlytotals].Expr1, Tbl_Customers.Customer"
    "_code\015\012FROM ((Tbl_Customers LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_co"
    "ntroller = Tbl_Users.ID) INNER JOIN [QueryCustomerOverdue>30daysOnlytotals] ON T"
    "bl_Customers.Customer_code = [QueryCustomerOverdue>30daysOnlytotals].Customer_co"
    "de) LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status."
    "ID\015\012WHERE (((Tbl_Customers.Credit_controller)=GetNumCreditController(fOSUs"
    "erName())) AND (([QueryCustomerOverdue>30daysOnlytotals].Expr1)>0))\015\012GROUP"
    " BY Tbl_Users.Name, Tbl_Customers.Name, Tbl_Customer_Status.Status, Tbl_Customer"
    "s.StatusDate, [QueryCustomerOverdue>30daysOnlytotals].Expr1, Tbl_Customers.Custo"
    "mer_code\015\012ORDER BY [QueryCustomerOverdue>30daysOnlytotals].Expr1 DESC;\015"
    "\012"
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
        dbText "Name" ="Tbl_Customers.StatusDate"
        dbInteger "ColumnWidth" ="1956"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days gone by"
        dbInteger "ColumnWidth" ="1372"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer name"
        dbInteger "ColumnWidth" ="4578"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[QueryCustomerOverdue>30daysOnlytotals].Expr1"
        dbText "Format" ="#,##0.00;-#,##0.00"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customer_Status.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
End
