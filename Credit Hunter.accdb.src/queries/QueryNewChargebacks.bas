﻿dbMemo "SQL" ="SELECT QueryChargebacksOpenNow.*, QueryChargebacksOpenPreviousDate.Customer_ID, "
    "Tbl_Customers.Name, QueryChargebacksOpenNow.Document_Number\015\012FROM Tbl_Cust"
    "omers INNER JOIN (QueryChargebacksOpenPreviousDate RIGHT JOIN QueryChargebacksOp"
    "enNow ON (QueryChargebacksOpenPreviousDate.Document_Number=QueryChargebacksOpenN"
    "ow.Document_Number) AND (QueryChargebacksOpenPreviousDate.Customer_ID=QueryCharg"
    "ebacksOpenNow.Customer_ID)) ON Tbl_Customers.Customer_code=QueryChargebacksOpenN"
    "ow.Customer_ID\015\012WHERE (((QueryChargebacksOpenPreviousDate.Customer_ID) Is "
    "Null))\015\012ORDER BY QueryChargebacksOpenNow.Document_Number;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="QueryChargebacksOpenPreviousDate.Customer_ID"
        dbInteger "ColumnWidth" ="4635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
