dbMemo "SQL" ="DELETE DISTINCTROW Tbl_Cash_Target.*, QueryCControllersInUseToday.Credit_control"
    "ler\015\012FROM Tbl_Cash_Target LEFT JOIN QueryCControllersInUseToday ON (Tbl_Ca"
    "sh_Target.Channel = QueryCControllersInUseToday.RetailOEM) AND (Tbl_Cash_Target."
    "CControllerID = QueryCControllersInUseToday.Credit_controller)\015\012WHERE (((Q"
    "ueryCControllersInUseToday.Credit_controller) Is Null));\015\012"
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
dbBoolean "UseTransaction" ="0"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Tbl_Cash_Target.CControllerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryCControllersInUseToday.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target.CashTargetInEUR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target.Channel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Channel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target.*"
        dbLong "AggregateType" ="-1"
    End
End
