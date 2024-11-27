dbMemo "SQL" ="SELECT Sum(([ExchangeRate]*[Original amount])) AS AmountInEUR, Tbl_CashCollected"
    ".CustomerID\015\012FROM Tbl_CashCollected INNER JOIN Tbl_Currencies ON Tbl_CashC"
    "ollected.Currency = Tbl_Currencies.CurrencyID\015\012WHERE (((Tbl_CashCollected."
    "[Payment Date])>=DateAdd(\"d\",1,DMax(\"[MonthEnd]\",\"[Tbl_MonthEnd]\",\"MonthE"
    "nd <#\" & Format(Date(),\"mm/dd/yy\") & \"#\"))))\015\012GROUP BY Tbl_CashCollec"
    "ted.CustomerID;\015\012"
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
dbMemo "Filter" ="([QueryTotalAlreadyCollectedInEUR].[CustomerID]=41485)"
Begin
    Begin
        dbText "Name" ="AmountInEUR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_CashCollected.CustomerID"
        dbLong "AggregateType" ="-1"
    End
End
