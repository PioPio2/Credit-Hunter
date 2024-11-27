UPDATE
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
  INNER JOIN Tbl_Invoices_History ON (
    QueryInvoicesPreviousDate.Customer_ID = Tbl_Invoices_History.Customer_ID
  )
  AND (
    QueryInvoicesPreviousDate.Date = Tbl_Invoices_History.Date
  )
  AND (
    QueryInvoicesPreviousDate.Document_Number = Tbl_Invoices_History.Document_Number
  )
SET
  Tbl_Invoices_History.[mEMO] = Date()
WHERE
  (
    (
      (
        QueryInvoicesLastDate.Tbl_Invoices.Update_date
      ) Is Null
    )
  );
