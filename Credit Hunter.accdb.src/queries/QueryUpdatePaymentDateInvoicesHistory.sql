UPDATE
  QueryInvoicesClosedInLastDate
  INNER JOIN Tbl_Invoices_History ON (
    QueryInvoicesClosedInLastDate.Customer_ID = Tbl_Invoices_History.Customer_ID
  )
  AND (
    QueryInvoicesClosedInLastDate.Document_Number = Tbl_Invoices_History.Document_Number
  )
  AND (
    QueryInvoicesClosedInLastDate.Date = Tbl_Invoices_History.Date
  )
SET
  Tbl_Invoices_History.PaymentDate = Date();
