dbMemo "SQL" ="DELETE Tbl_Invoices.Update_date\015\012FROM Tbl_Invoices\015\012WHERE (((Tbl_Inv"
    "oices.Update_date)<DateAdd(\"d\",-31,Now())));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Update_date"
        dbLong "AggregateType" ="-1"
    End
End
