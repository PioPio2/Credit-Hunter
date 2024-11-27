SELECT
  Tbl_Customers.Customer_code,
  Tbl_Invoices.Currency,
  Tbl_Currencies.ExchangeRate,
  Sum(
    IIf(
      [tbl_invoices].[update_date] = Date(),
      IIf(
        [tbl_invoices].[overdue_date] <= Date()-15,
        [Amount],
        0
      ),
      0
    )
  )* [ExchangeRate] AS Expr1
FROM
  (
    Tbl_Customers
    LEFT JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
  )
  LEFT JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID
GROUP BY
  Tbl_Customers.Customer_code,
  Tbl_Invoices.Currency,
  Tbl_Currencies.ExchangeRate;
