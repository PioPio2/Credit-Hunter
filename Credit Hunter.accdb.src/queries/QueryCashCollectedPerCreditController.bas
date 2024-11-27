dbMemo "SQL" ="SELECT Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Custome"
    "rs.Name, Query11.[AR Exposure in main currency], Query11.Update_date, Query11.[O"
    "verdue 90+ days], Query11.ExchangeRate, Tbl_Customer_Status.Status, Query11.[Tot"
    "al overdue on fiscal month end], Tbl_Customers.Country, Query11.[Overdue as of T"
    "oday in main currency], Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Ta"
    "rget in Main currency], IIf(IsNull([AmountInEUR]),0,[AmountInEUR]) AS [Already c"
    "ollected in Main currency], [MonthlyTargetInMainCurrency]-[Already collected in "
    "Main currency] AS [Still to be collected in Main currency], IIf([MonthlyTargetIn"
    "MainCurrency]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTargetIn"
    "MainCurrency],0) AS [% Cash Target Achieved]\015\012FROM ((Tbl_Customers LEFT JO"
    "IN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID) LEFT JO"
    "IN QueryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAl"
    "readyCollectedInEUR.CustomerID) LEFT JOIN Query11 ON Tbl_Customers.Customer_code"
    " = Query11.Customer_ID\015\012GROUP BY Tbl_Customers.Customer_code, Tbl_Customer"
    "s.Credit_controller, Tbl_Customers.Name, Query11.[AR Exposure in main currency],"
    " Query11.Update_date, Query11.[Overdue 90+ days], Query11.ExchangeRate, Tbl_Cust"
    "omer_Status.Status, Query11.[Total overdue on fiscal month end], Tbl_Customers.C"
    "ountry, Query11.[Overdue as of Today in main currency], Tbl_Customers.MonthlyTar"
    "getInMainCurrency, IIf(IsNull([AmountInEUR]),0,[AmountInEUR]), IIf([MonthlyTarge"
    "tInMainCurrency]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTarge"
    "tInMainCurrency],0)\015\012HAVING (((Tbl_Customers.Credit_controller)=1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Monthly Target in Main currency"
    End
    Begin
        dbText "Name" ="Already collected in Main currency"
    End
    Begin
        dbText "Name" ="Still to be collected in Main currency"
    End
    Begin
        dbText "Name" ="% Cash Target Achieved"
    End
End
