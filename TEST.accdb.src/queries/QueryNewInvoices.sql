INSERT INTO Tbl_Invoices_History (
  Customer_ID, [Date], Document_Number,
  Customer_reference, Type, Amount,
  Overdue_Date, [Currency], Update_date
)
SELECT
  QueryInvoicesLastDate.Customer_ID,
  QueryInvoicesLastDate.Date,
  QueryInvoicesLastDate.Document_Number,
  QueryInvoicesLastDate.Customer_reference,
  QueryInvoicesLastDate.Type,
  QueryInvoicesLastDate.Amount,
  QueryInvoicesLastDate.Overdue_Date,
  QueryInvoicesLastDate.Currency,
  Date() AS Expr1
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
