dbMemo "SQL" ="SELECT Tbl_DailyExposure.CustomerID, Max(Tbl_DailyExposure.ARExposure) AS MaxOfA"
    "RExposure\015\012FROM Tbl_DailyExposure\015\012GROUP BY Tbl_DailyExposure.Custom"
    "erID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Tbl_DailyExposure.CustomerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfARExposure"
        dbLong "AggregateType" ="-1"
    End
End
