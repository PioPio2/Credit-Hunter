dbMemo "SQL" ="SELECT Tbl_Users.Name, Tbl_Customers.Customer_code, Tbl_Customers.Name, Tbl_Invo"
    "ices.Date AS [Document date], Tbl_Invoices.Document_Number, Tbl_Invoices.Currenc"
    "y, Tbl_Invoices.Amount, Tbl_Invoices.Overdue_Date AS [Due date], Tbl_Customer_St"
    "atus.Status, Tbl_Invoices.Query, Tbl_Invoices.mEMO AS [Note]\015\012FROM ((Tbl_C"
    "ustomers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.C"
    "ustomer_ID) LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.I"
    "D) LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.I"
    "D\015\012WHERE (((Tbl_Invoices.Update_date)=Date() Or (Tbl_Invoices.Update_date)"
    "=Date()))\015\012GROUP BY Tbl_Users.Name, Tbl_Customers.Customer_code, Tbl_Custo"
    "mers.Name, Tbl_Invoices.Date, Tbl_Invoices.Document_Number, Tbl_Invoices.Currenc"
    "y, Tbl_Invoices.Amount, Tbl_Invoices.Overdue_Date, Tbl_Customer_Status.Status, T"
    "bl_Invoices.Query, Tbl_Invoices.mEMO\015\012HAVING (((Tbl_Invoices.Date)<=(DMin("
    "\"[MonthEnd]\",\"[Tbl_MonthEnd]\",\"[MonthEnd]>=#\" & Format(Date(),\"mm/dd/yy\""
    ") & \"#\"))-180)) OR (((Tbl_Invoices.Date)>(DMin(\"[MonthEnd]\",\"[Tbl_MonthEnd]"
    "\",\"[MonthEnd]>=#\" & Format(Date(),\"mm/dd/yy\") & \"#\"))-180) AND ((Tbl_Invo"
    "ices.Query)=19))\015\012ORDER BY Tbl_Users.Name, Tbl_Customers.Name;\015\012"
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
        dbText "Name" ="Tbl_Invoices.Document_Number"
        dbInteger "ColumnWidth" ="1902"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Amount"
        dbText "Format" ="Standard"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Document date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Due date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customer_Status.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Query"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Note"
        dbLong "AggregateType" ="-1"
    End
End
