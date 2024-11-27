Operation =1
Option =0
Begin InputTables
    Name ="Tbl_Customers"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Customers.Name"
    Alias ="Expr2"
    Expression ="Tbl_Customers.Address"
    Alias ="Expr3"
    Expression ="Tbl_Customers.Address2"
    Alias ="Expr4"
    Expression ="Tbl_Customers.Address3"
    Alias ="Expr5"
    Expression ="Tbl_Customers.Address4"
    Alias ="Expr1"
    Expression ="\"city\""
    Alias ="Expr6"
    Expression ="Tbl_Customers.Country"
    Alias ="Expr7"
    Expression ="Tbl_Customers.NextAppointment"
    Alias ="Expr8"
    Expression ="Tbl_Customers.StatusDate"
    Alias ="Expr9"
    Expression ="Tbl_Customers.Index"
    Alias ="Expr10"
    Expression ="Tbl_Customers.Note"
    Alias ="Expr11"
    Expression ="Tbl_Customers.TotalInsurance"
    Alias ="Expr12"
    Expression ="Tbl_Customers.Status"
    Alias ="Expr13"
    Expression ="Tbl_Customers.Timezone"
    Alias ="Expr14"
    Expression ="Tbl_Customers.Language"
    Alias ="Expr15"
    Expression ="Tbl_Customers.DSO"
    Alias ="Expr16"
    Expression ="Tbl_Customers.Area"
    Alias ="Expr17"
    Expression ="Tbl_Customers.LastStatementSent"
    Alias ="Expr2"
    Expression ="\"sales channel\""
    Alias ="Expr3"
    Expression ="\"active\""
    Alias ="Expr4"
    Expression ="\"total CL\""
    Alias ="Expr5"
    Expression ="\"sales person\""
    Alias ="Expr18"
    Expression ="Tbl_Customers.Customer_code"
    Alias ="Expr6"
    Expression ="\"postal code\""
    Alias ="Expr7"
    Expression ="\"county\""
    Alias ="Expr8"
    Expression ="\"group code\""
    Alias ="Expr9"
    Expression ="\"collection_category\""
    Alias ="Expr10"
    Expression ="\"collection_step\""
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
        dbText "Name" ="Tbl_Customers.Country"
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
        dbText "Name" ="Tbl_Customers.StatusDate"
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
        dbText "Name" ="Tbl_Customers.Address4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.NextAppointment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Index"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.TotalInsurance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Timezone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Language"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.DSO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.LastStatementSent"
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
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
    End
    Begin
        dbText "Name" ="Expr12"
    End
    Begin
        dbText "Name" ="Expr13"
    End
    Begin
        dbText "Name" ="Expr14"
    End
    Begin
        dbText "Name" ="Expr15"
    End
    Begin
        dbText "Name" ="Expr16"
    End
    Begin
        dbText "Name" ="Expr17"
    End
    Begin
        dbText "Name" ="Expr18"
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
    Bottom =591
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =289
        Bottom =626
        Top =0
        Name ="Tbl_Customers"
        Name =""
    End
End
