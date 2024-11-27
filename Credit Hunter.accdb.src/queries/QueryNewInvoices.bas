dbMemo "SQL" ="INSERT INTO Tbl_Invoices_History ( Customer_ID, [Date], Document_Number, Custome"
    "r_reference, Type, Amount, Overdue_Date, [Currency], Update_date )\015\012SELECT"
    " QueryInvoicesLastDate.Customer_ID, QueryInvoicesLastDate.Date, QueryInvoicesLas"
    "tDate.Document_Number, QueryInvoicesLastDate.Customer_reference, QueryInvoicesLa"
    "stDate.Type, QueryInvoicesLastDate.Amount, QueryInvoicesLastDate.Overdue_Date, Q"
    "ueryInvoicesLastDate.Currency, Date() AS Expr1\015\012FROM QueryInvoicesPrevious"
    "Date RIGHT JOIN QueryInvoicesLastDate ON (QueryInvoicesPreviousDate.Customer_ID="
    "QueryInvoicesLastDate.Customer_ID) AND (QueryInvoicesPreviousDate.Date=QueryInvo"
    "icesLastDate.Date) AND (QueryInvoicesPreviousDate.Document_Number=QueryInvoicesL"
    "astDate.Document_Number)\015\012WHERE (((QueryInvoicesLastDate.Document_Number)<"
    ">\" \") AND ((QueryInvoicesPreviousDate.Tbl_Invoices.Update_date) Is Null));\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="QueryInvoicesPreviousDate.Tbl_Invoices.Update_date"
        dbInteger "ColumnWidth" ="4800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
