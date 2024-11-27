SELECT
  Sum([Amount] * [ExchangeRate]) AS [AR Exposure in main currency],
  Tbl_Invoices.Update_date,
  Sum(
    IIf(
      [tbl_invoices.overdue_date] + 90 <= GetNextMonthEnd(),
      [amount],
      0
    )
  )* [ExchangeRate] AS [Overdue 90+ days],
  Tbl_Currencies.ExchangeRate,
  Sum(
    IIf(
      [tbl_invoices.overdue_date] <= GetNextMonthEnd(),
      [amount],
      0
    )
  )* [Tbl_Currencies].[ExchangeRate] AS [Total overdue on fiscal month end],
  Sum(
    IIf(
      [tbl_invoices.overdue_date] <= Now(),
      [amount],
      0
    )
  )* [Tbl_Currencies].[ExchangeRate] AS [Overdue as of Today in main currency],
  Tbl_Invoices.Customer_ID
FROM
  Tbl_Currencies
  RIGHT JOIN Tbl_Invoices ON Tbl_Currencies.CurrencyID = Tbl_Invoices.Currency
GROUP BY
  Tbl_Invoices.Update_date,
  Tbl_Currencies.ExchangeRate,
  Tbl_Invoices.Customer_ID
HAVING
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
  );
