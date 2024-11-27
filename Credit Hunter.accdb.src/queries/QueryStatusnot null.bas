dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Customer_Status.Status, Tbl_Customers.Na"
    "me\015\012FROM Tbl_Customers INNER JOIN Tbl_Customer_Status ON Tbl_Customers.Sta"
    "tus=Tbl_Customer_Status.ID\015\012WHERE (((Tbl_Customer_Status.Status) Is Not Nu"
    "ll) AND ((Tbl_Customers.Credit_controller)=GetNumCreditController(fOSUserName())"
    "))\015\012ORDER BY Tbl_Customer_Status.Status, Tbl_Customers.Name;\015\012"
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
        dbText "Name" ="Tbl_Customer_Status.Status"
        dbInteger "ColumnWidth" ="3532"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
