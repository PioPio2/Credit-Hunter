dbMemo "SQL" ="SELECT Tbl_Customers.Credit_controller, Tbl_Customers.RetailOEM\015\012FROM Tbl_"
    "Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices."
    "Customer_ID\015\012WHERE (((Tbl_Invoices.Update_date)=Date()))\015\012GROUP BY T"
    "bl_Customers.Credit_controller, Tbl_Customers.RetailOEM;\015\012"
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
        dbText "Name" ="Tbl_Customers.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.RetailOEM"
        dbLong "AggregateType" ="-1"
    End
End
