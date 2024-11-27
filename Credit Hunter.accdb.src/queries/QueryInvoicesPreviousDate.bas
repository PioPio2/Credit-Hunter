Operation =1
Option =0
Where ="(((Tbl_Invoices.Update_date)=#10/12/2024#))"
Begin InputTables
    Name ="Tbl_Invoices"
    Name ="Tbl_Types"
End
Begin OutputColumns
    Expression ="Tbl_Invoices.Update_date"
    Expression ="Tbl_Invoices.*"
    Expression ="Tbl_Types.Descripition"
End
Begin Joins
    LeftTable ="Tbl_Invoices"
    RightTable ="Tbl_Types"
    Expression ="Tbl_Invoices.Type = Tbl_Types.ID"
    Flag =2
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
        dbText "Name" ="Tbl_Invoices.Update_date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.SONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Overdue_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_reference"
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
        dbText "Name" ="Tbl_Invoices.Attachment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Query"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.mEMO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.QueryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.QueryToBePrinted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.CustomsInvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.PullTicketN#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Types.Descripition"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1391
    Bottom =930
    Left =-1
    Top =-1
    Right =1105
    Bottom =184
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
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Tbl_Types"
        Name =""
    End
End
