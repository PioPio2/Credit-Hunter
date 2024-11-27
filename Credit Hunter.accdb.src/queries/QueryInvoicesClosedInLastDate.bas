Operation =1
Option =0
Where ="(((QueryInvoicesLastDate.Date) Is Null))"
Begin InputTables
    Name ="QueryInvoicesPreviousDate"
    Name ="QueryInvoicesLastDate"
    Name ="Tbl_queries"
End
Begin OutputColumns
    Expression ="QueryInvoicesPreviousDate.*"
    Expression ="Tbl_queries.Query"
End
Begin Joins
    LeftTable ="QueryInvoicesPreviousDate"
    RightTable ="QueryInvoicesLastDate"
    Expression ="QueryInvoicesPreviousDate.Customer_ID = QueryInvoicesLastDate.Customer_ID"
    Flag =2
    LeftTable ="QueryInvoicesPreviousDate"
    RightTable ="QueryInvoicesLastDate"
    Expression ="QueryInvoicesPreviousDate.Date = QueryInvoicesLastDate.Date"
    Flag =2
    LeftTable ="QueryInvoicesPreviousDate"
    RightTable ="QueryInvoicesLastDate"
    Expression ="QueryInvoicesPreviousDate.Document_Number = QueryInvoicesLastDate.Document_Numbe"
        "r"
    Flag =2
    LeftTable ="QueryInvoicesPreviousDate"
    RightTable ="Tbl_queries"
    Expression ="QueryInvoicesPreviousDate.Query = Tbl_queries.ID"
    Flag =2
End
Begin OrderBy
    Expression ="QueryInvoicesLastDate.Date"
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
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Update_date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_queries.Query"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Document_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Overdue_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Customer_reference"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.SONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.OriginalAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.PullTicketN#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Query"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.mEMO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.QueryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.QueryToBePrinted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Attachment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.CustomsInvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Types.Descripition"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1121
    Bottom =930
    Left =-1
    Top =-1
    Right =1105
    Bottom =253
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="QueryInvoicesPreviousDate"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="QueryInvoicesLastDate"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Tbl_queries"
        Name =""
    End
End
