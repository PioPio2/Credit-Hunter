SELECT
  QueryInvoicesPreviousDate.*,
  Tbl_queries.Query
FROM
  (
    QueryInvoicesPreviousDate
    LEFT JOIN QueryInvoicesLastDate ON (
      QueryInvoicesPreviousDate.Document_Number = QueryInvoicesLastDate.Document_Number
    )
    AND (
      QueryInvoicesPreviousDate.Date = QueryInvoicesLastDate.Date
    )
    AND (
      QueryInvoicesPreviousDate.Customer_ID = QueryInvoicesLastDate.Customer_ID
    )
  )
  LEFT JOIN Tbl_queries ON QueryInvoicesPreviousDate.Query = Tbl_queries.ID
WHERE
  (
    (
      (QueryInvoicesLastDate.Date) Is Null
    )
  )
ORDER BY
  QueryInvoicesLastDate.Date;
