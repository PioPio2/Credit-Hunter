DELETE QueryInvoicesLastDate.Customer_ID,
QueryInvoicesLastDate.Date,
QueryInvoicesLastDate.Document_Number,
QueryInvoicesLastDate.Customer_reference,
QueryInvoicesLastDate.Type,
QueryInvoicesLastDate.Amount,
QueryInvoicesLastDate.Overdue_Date,
QueryInvoicesLastDate.Currency,
QueryInvoicesLastDate.Query,
QueryInvoicesLastDate.mEMO,
QueryInvoicesPreviousDate.Tbl_Invoices.Update_date
FROM
  QueryInvoicesPreviousDate
  RIGHT JOIN QueryInvoicesLastDate ON (
    QueryInvoicesPreviousDate.Customer_ID = QueryInvoicesLastDate.Customer_ID
  )
  AND (
    QueryInvoicesPreviousDate.Date = QueryInvoicesLastDate.Date
  )
  AND (
    QueryInvoicesPreviousDate.Document_Number = QueryInvoicesLastDate.Document_Number
  )
WHERE
  (
    (
      (
        QueryInvoicesLastDate.Document_Number
      )<> " "
    )
    AND (
      (
        QueryInvoicesPreviousDate.Tbl_Invoices.Update_date
      ) Is Null
    )
  );
