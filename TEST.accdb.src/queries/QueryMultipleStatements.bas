Operation =1
Option =0
Where ="(                 ( (Tbl_Customers.Name) Like '*a*' OR (Tbl_Customers.Name) Like"
    " '*c*' OR (Tbl_Customers.Name) Like '*b*' )                AND (                "
    "         (                                 Tbl_Invoices.Update_date)=Date()))"
Begin InputTables
    Name ="Tbl_Customers"
    Name ="Tbl_Invoices"
End
Begin OutputColumns
    Expression ="Tbl_Customers.Name"
    Expression ="Tbl_Invoices.Date"
    Expression ="Tbl_Invoices.Document_Number"
    Expression ="Tbl_Invoices.Amount"
    Expression ="Tbl_Invoices.Overdue_Date"
    Expression ="Tbl_Invoices.mEMO"
End
Begin Joins
    LeftTable ="Tbl_Customers"
    RightTable ="Tbl_Invoices"
    Expression ="Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID"
    Flag =1
End
Begin OrderBy
    Expression ="Tbl_Customers.Name"
    Flag =0
    Expression ="Tbl_Invoices.Date"
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
        dbText "Name" ="Tbl_Invoices.Overdue_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Document_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.mEMO"
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
    Bottom =529
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =98
        Top =17
        Right =242
        Bottom =161
        Top =0
        Name ="Tbl_Customers"
        Name =""
    End
    Begin
        Left =506
        Top =35
        Right =650
        Bottom =525
        Top =0
        Name ="Tbl_Invoices"
        Name =""
    End
End
