SELECT
  DISTINCTROW Tbl_Invoices.Customer_ID,
  Sum(
    ([Amount] * [ExchangeRate])
  ) AS TotalExposureConvertedInEUR
FROM
  (
    Tbl_Customers
    INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
  )
  INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
GROUP BY
  Tbl_Invoices.Customer_ID;
