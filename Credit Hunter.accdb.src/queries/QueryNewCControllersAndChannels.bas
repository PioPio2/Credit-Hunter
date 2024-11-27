dbMemo "SQL" ="INSERT INTO Tbl_Cash_Target ( Channel, CControllerID )\015\012SELECT QueryCContr"
    "ollersInUseToday.RetailOEM, Tbl_Users.ID\015\012FROM (QueryCControllersInUseToda"
    "y LEFT JOIN Tbl_Cash_Target ON (QueryCControllersInUseToday.Credit_controller = "
    "Tbl_Cash_Target.CControllerID) AND (QueryCControllersInUseToday.RetailOEM = Tbl_"
    "Cash_Target.Channel)) LEFT JOIN Tbl_Users ON QueryCControllersInUseToday.Credit_"
    "controller = Tbl_Users.ID\015\012WHERE (((Tbl_Cash_Target.CashTargetInEUR) Is Nu"
    "ll));\015\012"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="QueryCControllersInUseToday.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryCControllersInUseToday.RetailOEM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target.CashTargetInEUR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.UserName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Users.ID"
        dbLong "AggregateType" ="-1"
    End
End
