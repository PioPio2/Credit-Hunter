﻿dbMemo "SQL" ="UPDATE QueryInvoicesClosedInLastDate INNER JOIN Tbl_Invoices_History ON (QueryIn"
    "voicesClosedInLastDate.Customer_ID=Tbl_Invoices_History.Customer_ID) AND (QueryI"
    "nvoicesClosedInLastDate.Document_Number=Tbl_Invoices_History.Document_Number) AN"
    "D (QueryInvoicesClosedInLastDate.Date=Tbl_Invoices_History.Date) SET Tbl_Invoice"
    "s_History.PaymentDate = Date();\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End