dbMemo "SQL" ="SELECT Tbl_Users.Name, Tbl_Customers.Customer_code, Tbl_Customers.Name AS [Custo"
    "mer name], Tbl_Customers.Status, Tbl_Customers.StatusDate\015\012FROM Tbl_Custom"
    "ers LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller=Tbl_Users.ID\015\012W"
    "HERE (((Tbl_Customers.Credit_controller)=GetNumCreditController(fOSUserName())))"
    "\015\012GROUP BY Tbl_Users.Name, Tbl_Customers.Customer_code, Tbl_Customers.Name"
    ", Tbl_Customers.Status, Tbl_Customers.StatusDate\015\012HAVING (((Tbl_Customers."
    "Status) Is Not Null))\015\012ORDER BY Tbl_Customers.Name;\015\012"
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
        dbText "Name" ="Tbl_Customers.Status"
        dbInteger "ColumnWidth" ="3369"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.StatusDate"
        dbInteger "ColumnWidth" ="1956"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer name"
        dbInteger "ColumnWidth" ="3804"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbInteger "ColumnWidth" ="1589"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
