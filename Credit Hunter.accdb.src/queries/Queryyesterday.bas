Operation =1
Option =0
Where ="(((Tbl_Invoices.Update_date)=#3/31/2015#) And ((Tbl_Invoices.QueryToBePrinted)=N"
    "o))"
Begin InputTables
    Name ="Tbl_Invoices"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Invoices.Customer_ID"
    Alias ="Expr2"
    Expression ="Tbl_Invoices.Update_date"
    Alias ="Expr3"
    Expression ="Tbl_Invoices.Date"
    Alias ="Expr4"
    Expression ="Tbl_Invoices.Document_Number"
    Alias ="Expr5"
    Expression ="Tbl_Invoices.QueryToBePrinted"
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
    Bottom =592
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Tbl_Invoices"
        Name =""
    End
End
