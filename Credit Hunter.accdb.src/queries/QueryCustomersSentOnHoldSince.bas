dbMemo "SQL" ="SELECT Tbl_Customers.Name, Tbl_Customers.Customer_code, TblNotes.CustomerCode, T"
    "blNotes.Note\015\012FROM Tbl_Customers LEFT JOIN TblNotes ON Tbl_Customers.Custo"
    "mer_code = TblNotes.CustomerCode\015\012WHERE (((TblNotes.ID)>55411))\015\012GRO"
    "UP BY Tbl_Customers.Name, Tbl_Customers.Customer_code, TblNotes.CustomerCode, Tb"
    "lNotes.Note\015\012HAVING (((TblNotes.Note) Like \"*Modify status into: HOLD*\")"
    ");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "ReplicableBool" ="-1"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="1845"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TblNotes.CustomerCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TblNotes.Note"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
End
