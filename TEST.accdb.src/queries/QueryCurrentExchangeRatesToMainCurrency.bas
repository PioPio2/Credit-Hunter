dbMemo "SQL" ="SELECT Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.Fis"
    "calYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.Exchan"
    "geRateToMainCurrency, Tbl_MonthEnd.MonthEnd\015\012FROM Tbl_MonthEnd LEFT JOIN T"
    "bl_Cash_Target_Breakdown ON (Tbl_MonthEnd.FiscalYear = Tbl_Cash_Target_Breakdown"
    ".FiscalYear) AND (Tbl_MonthEnd.FiscalQuarter = Tbl_Cash_Target_Breakdown.FiscalQ"
    "uarter) AND (Tbl_MonthEnd.FiscalMonth = Tbl_Cash_Target_Breakdown.FiscalMonth)\015"
    "\012GROUP BY Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdo"
    "wn.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown."
    "ExchangeRateToMainCurrency, Tbl_MonthEnd.MonthEnd\015\012HAVING (((Tbl_MonthEnd."
    "MonthEnd)=DMin(\"MonthEnd\",\"Tbl_MonthEnd\",\"MonthEnd>=date()\")));\015\012"
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
        dbText "Name" ="Tbl_Cash_Target_Breakdown.OriginalCurrency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target_Breakdown.FiscalYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target_Breakdown.FiscalMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tbl_MonthEnd.MonthEnd"
        dbLong "AggregateType" ="-1"
    End
End
