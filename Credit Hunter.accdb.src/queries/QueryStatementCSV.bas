Operation =1
Option =0
Where ="((([Tbl_Invoices].[Update_date])=Date()))"
Begin InputTables
    Name ="Tbl_Invoices"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="Tbl_Invoices.Customer_ID"
    Alias ="updatedate"
    Expression ="Format([Update_date],\"dd/mm/yyyy\"\" 00\"\":\"\"00\"\":\"\"00\"\"\")"
    Alias ="invoicedate"
    Expression ="Format([Date],\"dd/mm/yyyy\"\" 00\"\":\"\"00\"\":\"\"00\"\"\")"
    Alias ="Expr2"
    Expression ="Tbl_Invoices.Document_Number"
    Alias ="Expr3"
    Expression ="Tbl_Invoices.Customer_reference"
    Alias ="Expr4"
    Expression ="Tbl_Invoices.SONumber"
    Alias ="Expr5"
    Expression ="Tbl_Invoices.Type"
    Alias ="Expr6"
    Expression ="Tbl_Invoices.OriginalAmount"
    Alias ="Expr7"
    Expression ="Tbl_Invoices.Amount"
    Alias ="invoiceduedate"
    Expression ="Format([Overdue_Date],\"dd/mm/yyyy\"\" 00\"\":\"\"00\"\":\"\"00\"\"\")"
    Alias ="Expr8"
    Expression ="Tbl_Invoices.Currency"
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
        dbText "Name" ="Tbl_Invoices.Document_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_reference"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.SONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.OriginalAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="updatedate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="invoicedate"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="invoiceduedate"
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
    Bottom =506
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =429
        Bottom =525
        Top =0
        Name ="Tbl_Invoices"
        Name =""
    End
End
