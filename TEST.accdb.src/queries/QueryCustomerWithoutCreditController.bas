Operation =1
Option =0
Where ="((([Tbl_Customers].[Credit_controller]) Is Null))"
Begin InputTables
    Name ="Tbl_Customers"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Customers.Credit_controller"
    Alias ="Expr2"
    Expression ="Tbl_Customers.Customer_code"
    Alias ="Expr3"
    Expression ="Tbl_Customers.Name"
    Alias ="Expr4"
    Expression ="Tbl_Customers.Country"
End
Begin OrderBy
    Expression ="Tbl_Customers.Name"
    Flag =0
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
        dbText "Name" ="Tbl_Customers.Credit_controller"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "DisplayControl" ="110"
        dbText "RowSourceType" ="Table/Query"
        dbMemo "RowSource" ="SELECT Tbl_Users.ID, Tbl_Users.UserName FROM Tbl_Users; "
        dbInteger "ColumnCount" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
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
    Bottom =298
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =241
        Bottom =248
        Top =0
        Name ="Tbl_Customers"
        Name =""
    End
End
