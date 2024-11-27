Operation =5
Option =0
Where ="((([Tbl_Invoices_History].[Date])<DateAdd(\"yyyy\",-1,Now())))"
Begin InputTables
    Name ="Tbl_Invoices_History"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Invoices_History.Date"
    Alias ="Expr2"
    Expression ="[Tbl_Invoices_History].[Date]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
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
    Bottom =424
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =335
        Top =6
        Right =813
        Bottom =333
        Top =0
        Name ="Tbl_Invoices_History"
        Name =""
    End
End
