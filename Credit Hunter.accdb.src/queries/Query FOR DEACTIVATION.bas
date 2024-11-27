dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Customers.Name, Tbl_Customers.Country, T"
    "bl_Areas.Area, Tbl_Users.Name\015\012FROM (Tbl_Customers LEFT JOIN Tbl_Areas ON "
    "Tbl_Customers.Area = Tbl_Areas.ID) LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_c"
    "ontroller = Tbl_Users.ID;\015\012"
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
        dbText "Name" ="Tbl_Areas.Area"
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
        dbText "Name" ="Tbl_Customers.Country"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.Name"
        dbLong "AggregateType" ="-1"
    End
End
