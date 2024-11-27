dbMemo "SQL" ="SELECT Tbl_Customers.Credit_controller, Tbl_Customers.Customer_code, Tbl_Custome"
    "rs.Name, Sum([QueryCustomerOverdue>30daysOnlytotals].Expr1) AS aa, Tbl_Customer_"
    "Status.Status\015\012FROM (Tbl_Customers INNER JOIN [QueryCustomerOverdue>30days"
    "Onlytotals] ON Tbl_Customers.Customer_code = [QueryCustomerOverdue>30daysOnlytot"
    "als].Customer_code) INNER JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl"
    "_Customer_Status.ID\015\012GROUP BY Tbl_Customers.Credit_controller, Tbl_Custome"
    "rs.Customer_code, Tbl_Customers.Name, Tbl_Customer_Status.Status\015\012HAVING ("
    "((Tbl_Customers.Credit_controller)=GetNumCreditController(fOSUserName())) AND (("
    "Sum([QueryCustomerOverdue>30daysOnlytotals].Expr1))<=0 Or (Sum([QueryCustomerOve"
    "rdue>30daysOnlytotals].Expr1)) Is Null) AND ((Tbl_Customer_Status.Status) Is Not"
    " Null))\015\012ORDER BY Tbl_Customer_Status.Status;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="aa"
        dbText "Format" ="$#,##0.00;-$#,##0.00"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Credit_controller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Customer_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customer_Status.Status"
        dbLong "AggregateType" ="-1"
    End
End
