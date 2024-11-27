SELECT
  DISTINCTROW Tbl_Invoices_History.Customer_ID,
  Sum([Amount] * [ExchangeRate]) AS TotalInvoicedConvertedInEUR
FROM
  Tbl_Currencies
  INNER JOIN (
    Tbl_Customers
    INNER JOIN Tbl_Invoices_History ON Tbl_Customers.Customer_code = Tbl_Invoices_History.Customer_ID
  ) ON Tbl_Currencies.CurrencyID = Tbl_Invoices_History.Currency
GROUP BY
  Tbl_Invoices_History.Customer_ID;
