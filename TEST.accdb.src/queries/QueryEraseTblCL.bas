Operation =5
Option =0
Begin InputTables
    Name ="Tbl_CL"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_CL.Customer_code"
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
    Bottom =111
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =122
        Top =0
        Name ="Tbl_CL"
        Name =""
    End
End
