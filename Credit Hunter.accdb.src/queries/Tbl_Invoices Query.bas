Operation =1
Option =0
Begin InputTables
    Name ="Tbl_Invoices"
End
Begin OutputColumns
    Alias ="Espr1"
    Expression ="Tbl_Invoices.Customer_ID"
    Alias ="Espr2"
    Expression ="Tbl_Invoices.Update_date"
    Alias ="Espr3"
    Expression ="Tbl_Invoices.Document_Number"
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
End
Begin
    State =0
    Left =0
    Top =0
    Right =1391
    Bottom =808
    Left =-1
    Top =-1
    Right =1375
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =88
        Top =0
        Name ="Tbl_Invoices"
        Name =""
    End
End
