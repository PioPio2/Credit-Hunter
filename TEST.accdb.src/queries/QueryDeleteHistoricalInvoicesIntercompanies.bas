﻿dbMemo "SQL" ="DELETE Tbl_Invoices_History.*, Tbl_Users.UserName\015\012FROM Tbl_Users INNER JO"
    "IN (Tbl_Customers INNER JOIN Tbl_Invoices_History ON Tbl_Customers.Customer_code"
    "=Tbl_Invoices_History.Customer_ID) ON Tbl_Users.ID=Tbl_Customers.Credit_controll"
    "er\015\012WHERE (((Tbl_Users.UserName)=\"int\"));\015\012"
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
