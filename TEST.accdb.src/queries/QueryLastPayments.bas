Operation =1
Option =0
Begin InputTables
    Name ="Tbl_Customers"
End
Begin OutputColumns
    Alias ="Espr1"
    Expression ="Tbl_Customers.Customer_code"
    Alias ="Espr2"
    Expression ="Tbl_Customers.Name"
    Alias ="Espr3"
    Expression ="Tbl_Customers.ToSendStatement"
    Alias ="Espr4"
    Expression ="Tbl_Customers.Credit_controller"
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
        dbText "Name" ="Espr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr4"
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
    Bottom =440
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =239
        Bottom =47
        Top =0
        Name ="Tbl_Customers"
        Name =""
    End
End
