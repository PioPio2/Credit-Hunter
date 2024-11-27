SELECT
  Tbl_Invoices.Customer_ID,
  Sum(
    [Amount] / [ExchangeRateToMainCurrency]
  ) AS TotalOverdueOver90
FROM
  Tbl_Invoices
  LEFT JOIN QueryCurrentExchangeRatesToMainCurrency ON Tbl_Invoices.Currency = QueryCurrentExchangeRatesToMainCurrency.OriginalCurrency
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
    AND (
      (Tbl_Invoices.Overdue_Date)<= DateAdd('d',-90, [MonthEnd])
    )
  )
GROUP BY
  Tbl_Invoices.Customer_ID
ORDER BY
  Tbl_Invoices.Customer_ID;
