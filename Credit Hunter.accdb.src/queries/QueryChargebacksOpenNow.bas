dbMemo "SQL" ="SELECT Tbl_Invoices.*, Tbl_Types.ToFillChargbackFile, Tbl_Invoices.Update_date\015"
    "\012FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type=Tbl_Types.ID\015"
    "\012WHERE (((Tbl_Types.ToFillChargbackFile)=True) And ((Tbl_Invoices.Update_date"
    ")=#4/7/2010#));\015\012"
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
        dbText "Name" ="Tbl_Invoices.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Update_date"
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
        dbText "Name" ="Tbl_Invoices.Customer_reference"
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
        dbText "Name" ="Tbl_Invoices.Overdue_Date"
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
        dbText "Name" ="Tbl_Types.ToFillChargbackFile"
        dbLong "AggregateType" ="-1"
    End
End
