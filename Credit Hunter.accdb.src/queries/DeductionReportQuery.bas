dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Customers.Name, Tbl_Invoices.Type\015\012"
    "FROM Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_"
    "Invoices.Customer_ID\015\012WHERE (((Tbl_Invoices.Update_date)=Date()))\015\012G"
    "ROUP BY Tbl_Customers.Customer_code, Tbl_Customers.Name, Tbl_Invoices.Type\015\012"
    "HAVING (((Tbl_Invoices.Type)=4))\015\012ORDER BY Tbl_Customers.Name;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "ReplicableBool" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Tbl_Invoices.Type"
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
End
