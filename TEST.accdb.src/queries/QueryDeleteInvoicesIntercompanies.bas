﻿dbMemo "SQL" ="DELETE Tbl_Users.UserName, Tbl_Invoices.*\015\012FROM (Tbl_Users INNER JOIN Tbl_"
    "Customers ON Tbl_Users.ID=Tbl_Customers.Credit_controller) INNER JOIN Tbl_Invoic"
    "es ON Tbl_Customers.Customer_code=Tbl_Invoices.Customer_ID\015\012WHERE (((Tbl_U"
    "sers.UserName)=\"int\"));\015\012"
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