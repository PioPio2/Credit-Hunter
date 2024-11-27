SELECT
  Tbl_Invoices.Customer_ID,
  Sum(
    [Amount] / [ExchangeRateToMainCurrency]
  ) AS ARExposure
FROM
  Tbl_Invoices
  LEFT JOIN QueryCurrentExchangeRatesToMainCurrency ON Tbl_Invoices.Currency = QueryCurrentExchangeRatesToMainCurrency.OriginalCurrency
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
GROUP BY
  Tbl_Invoices.Customer_ID
ORDER BY
  Tbl_Invoices.Customer_ID;
