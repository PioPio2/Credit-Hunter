Operation =1
Option =0
Begin InputTables
    Name ="Tbl_CashCollected"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_CashCollected.CustomerID"
    Alias ="a"
    Expression ="Format([Payment Date],\"dd/mm/yyyy\"\" 00\"\":\"\"00\"\":\"\"00\"\"\")"
    Alias ="Expr2"
    Expression ="Tbl_CashCollected.Currency"
    Alias ="Expr3"
    Expression ="Tbl_CashCollected.Amount"
    Alias ="Expr4"
    Expression ="Tbl_CashCollected.[Original amount]"
    Alias ="Expr5"
    Expression ="Tbl_CashCollected.RETnumber"
    Alias ="Expr6"
    Expression ="Tbl_CashCollected.FiscalYear"
    Alias ="Expr7"
    Expression ="Tbl_CashCollected.FiscalMonth"
    Alias ="Expr8"
    Expression ="Tbl_CashCollected.FiscalQuarter"
End
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
        dbText "Name" ="Tbl_CashCollected.CustomerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.[Original amount]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.RETnumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.FiscalYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.FiscalMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.FiscalQuarter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="a"
        dbInteger "ColumnWidth" ="4155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
    End
    Begin
        dbText "Name" ="Expr2"
    End
    Begin
        dbText "Name" ="Expr3"
    End
    Begin
        dbText "Name" ="Expr4"
    End
    Begin
        dbText "Name" ="Expr5"
    End
    Begin
        dbText "Name" ="Expr6"
    End
    Begin
        dbText "Name" ="Expr7"
    End
    Begin
        dbText "Name" ="Expr8"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =837
    Left =-1
    Top =-1
    Right =1689
    Bottom =489
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =300
        Bottom =406
        Top =0
        Name ="Tbl_CashCollected"
        Name =""
    End
End
