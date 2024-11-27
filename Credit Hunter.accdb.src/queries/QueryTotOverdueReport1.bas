Operation =1
Option =0
Where ="((([Tbl_Historical_Statements].[Overdue_Date])<=[Update_date]))"
Begin InputTables
    Name ="Tbl_Historical_Statements"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Historical_Statements.Update_date"
    Alias ="Expr2"
    Expression ="Tbl_Historical_Statements.Customer_ID"
    Alias ="SumOfAmount"
    Expression ="Sum(Tbl_Historical_Statements.Amount)"
End
Begin Groups
    Expression ="Tbl_Historical_Statements.Update_date"
    GroupLevel =0
    Expression ="Tbl_Historical_Statements.Customer_ID"
    GroupLevel =0
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfAmount"
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
    Bottom =285
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =42
        Top =6
        Right =212
        Bottom =248
        Top =0
        Name ="Tbl_Historical_Statements"
        Name =""
    End
End
