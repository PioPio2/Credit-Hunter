SELECT
  Tbl_Cash_Target_Breakdown.OriginalCurrency,
  Tbl_Cash_Target_Breakdown.FiscalYear,
  Tbl_Cash_Target_Breakdown.FiscalMonth,
  Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency,
  Tbl_MonthEnd.MonthEnd
FROM
  Tbl_MonthEnd
  LEFT JOIN Tbl_Cash_Target_Breakdown ON (
    Tbl_MonthEnd.FiscalYear = Tbl_Cash_Target_Breakdown.FiscalYear
  )
  AND (
    Tbl_MonthEnd.FiscalQuarter = Tbl_Cash_Target_Breakdown.FiscalQuarter
  )
  AND (
    Tbl_MonthEnd.FiscalMonth = Tbl_Cash_Target_Breakdown.FiscalMonth
  )
GROUP BY
  Tbl_Cash_Target_Breakdown.OriginalCurrency,
  Tbl_Cash_Target_Breakdown.FiscalYear,
  Tbl_Cash_Target_Breakdown.FiscalMonth,
  Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency,
  Tbl_MonthEnd.MonthEnd
HAVING
  (
    (
      (Tbl_MonthEnd.MonthEnd)= DMin(
        "MonthEnd", "Tbl_MonthEnd", "MonthEnd>=date()"
      )
    )
  );
