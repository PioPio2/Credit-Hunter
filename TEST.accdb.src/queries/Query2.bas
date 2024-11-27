dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Custome"
    "rs.Name, Query1.[AR Exposure in main currency], Query1.Update_date, Query1.[Over"
    "due 90+ days], Query1.ExchangeRate, Tbl_Customer_Status.Status, Query1.[Total ov"
    "erdue on fiscal month end], Tbl_Customers.Country, Query1.[Overdue as of Today i"
    "n main currency], Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Target i"
    "n Main currency], IIf(IsNull([AmountInEUR]),0,[AmountInEUR]) AS [Already collect"
    "ed in Main currency], [MonthlyTargetInMainCurrency]-[Already collected in Main c"
    "urrency] AS [Still to be collected in Main currency], IIf([MonthlyTargetInMainCu"
    "rrency]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTargetInMainCu"
    "rrency],0) AS [% Cash Target Achieved]\015\012FROM ((Tbl_Customers LEFT JOIN Tbl"
    "_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID) LEFT JOIN Que"
    "ryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyC"
    "ollectedInEUR.CustomerID) LEFT JOIN Query1 ON Tbl_Customers.Customer_code = Quer"
    "y1.Customer_ID\015\012GROUP BY Tbl_Customers.Customer_code, Tbl_Customers.Credit"
    "_controller, Tbl_Customers.Name, Query1.[AR Exposure in main currency], Query1.U"
    "pdate_date, Query1.[Overdue 90+ days], Query1.ExchangeRate, Tbl_Customer_Status."
    "Status, Query1.[Total overdue on fiscal month end], Tbl_Customers.Country, Query"
    "1.[Overdue as of Today in main currency], Tbl_Customers.MonthlyTargetInMainCurre"
    "ncy, IIf(IsNull([AmountInEUR]),0,[AmountInEUR]), IIf([MonthlyTargetInMainCurrenc"
    "y]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTargetInMainCurrenc"
    "y],0)\015\012HAVING (((Tbl_Customers.Credit_controller)=1));\015\012"
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
        dbText "Name" ="Tbl_Customer_Status.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Customers.Country"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.[AR Exposure in main currency]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Monthly Target in Main currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.Update_date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.[Overdue 90+ days]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.[Total overdue on fiscal month end]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.[Overdue as of Today in main currency]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Query1.ExchangeRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Already collected in Main currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Still to be collected in Main currency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="% Cash Target Achieved"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
    End
End
