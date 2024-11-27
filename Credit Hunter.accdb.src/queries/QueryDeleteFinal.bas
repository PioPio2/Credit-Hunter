dbMemo "SQL" ="DELETE Query6.Update_date, Tbl_InvoiceAttachments.*\015\012FROM Tbl_InvoiceAttac"
    "hments LEFT JOIN Query6 ON Tbl_InvoiceAttachments.DocumentID = Query6.a\015\012W"
    "HERE (((Query6.Update_date) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "ReplicableBool" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Query6.Update_date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Attachment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Document_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Customer_reference"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query55.a"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_InvoiceAttachments.DocumentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_InvoiceAttachments.AttachName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_InvoiceAttachments.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_InvoiceAttachments.CustomerID"
        dbLong "AggregateType" ="-1"
    End
End
