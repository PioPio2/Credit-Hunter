dbMemo "SQL" ="SELECT Tbl_Customers.Credit_controller, Tbl_Customers.Country, Tbl_Customers.Cus"
    "tomer_code, Tbl_Customers.Name, Tbl_Invoices.Date, Tbl_Invoices.Document_Number,"
    " Tbl_Invoices.Type, Tbl_Invoices.Amount, Tbl_Invoices.Currency, Tbl_Invoices.Que"
    "ry, Tbl_Invoices.mEMO\015\012FROM Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_C"
    "ustomers.Customer_code=Tbl_Invoices.Customer_ID\015\012WHERE (((Tbl_Customers.Cr"
    "edit_controller)=GetNumCreditController(fOSUserName())) AND ((Tbl_Invoices.Type)"
    "=4 Or (Tbl_Invoices.Type)=2) AND ((Tbl_Invoices.Update_date)=Date()))\015\012ORD"
    "ER BY Tbl_Customers.Country, Tbl_Customers.Name;\015\012"
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
        dbText "Name" ="Tbl_Invoices.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Document_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Amount"
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
        dbText "Name" ="Tbl_Customers.Country"
        dbLong "AggregateType" ="-1"
    End
End
