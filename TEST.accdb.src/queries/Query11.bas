dbMemo "SQL" ="SELECT Sum([Amount]*[ExchangeRate]) AS [AR Exposure in main currency], Tbl_Invoi"
    "ces.Update_date, Sum(IIf([tbl_invoices.overdue_date]+90<=GetNextMonthEnd(),[amou"
    "nt],0))*[ExchangeRate] AS [Overdue 90+ days], Tbl_Currencies.ExchangeRate, Sum(I"
    "If([tbl_invoices.overdue_date]<=GetNextMonthEnd(),[amount],0))*[Tbl_Currencies]."
    "[ExchangeRate] AS [Total overdue on fiscal month end], Sum(IIf([tbl_invoices.ove"
    "rdue_date]<=Now(),[amount],0))*[Tbl_Currencies].[ExchangeRate] AS [Overdue as of"
    " Today in main currency], Tbl_Invoices.Customer_ID\015\012FROM Tbl_Currencies RI"
    "GHT JOIN Tbl_Invoices ON Tbl_Currencies.CurrencyID = Tbl_Invoices.Currency\015\012"
    "GROUP BY Tbl_Invoices.Update_date, Tbl_Currencies.ExchangeRate, Tbl_Invoices.Cus"
    "tomer_ID\015\012HAVING (((Tbl_Invoices.Update_date)=Date()));\015\012"
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
        dbText "Name" ="AR Exposure in main currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Update_date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Overdue 90+ days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Currencies.ExchangeRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Total overdue on fiscal month end"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Overdue as of Today in main currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Invoices.Customer_ID"
        dbLong "AggregateType" ="-1"
    End
End
