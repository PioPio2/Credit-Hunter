Operation =1
Option =0
Begin InputTables
    Name ="Tbl_Customers"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Customers.Customer_code"
    Alias ="Expr2"
    Expression ="Tbl_Customers.Credit_controller"
    Alias ="Expr3"
    Expression ="Tbl_Customers.Name"
    Alias ="Expr4"
    Expression ="Tbl_Customers.Address"
    Alias ="Expr5"
    Expression ="Tbl_Customers.Address2"
    Alias ="Expr6"
    Expression ="Tbl_Customers.Address3"
    Alias ="City"
    Expression ="\"\""
    Alias ="Expr7"
    Expression ="Tbl_Customers.Country"
    Alias ="Expr8"
    Expression ="Tbl_Customers.Area"
    Alias ="Expr9"
    Expression ="Tbl_Customers.RetailOEM"
    Alias ="totCL"
    Expression ="0"
    Alias ="insuredCL"
    Expression ="0"
    Alias ="Expr10"
    Expression ="Tbl_Customers.Language"
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
        dbText "Name" ="totCL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Address"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Address3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Country"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.RetailOEM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="insuredCL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Language"
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
    Begin
        dbText "Name" ="Expr9"
    End
    Begin
        dbText "Name" ="Expr10"
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
    Bottom =608
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =345
        Bottom =610
        Top =0
        Name ="Tbl_Customers"
        Name =""
    End
End
