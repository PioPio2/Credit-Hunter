SELECT
  Tbl_Invoices.Customer_ID AS Expr1,
  Format(
    [Update_date], "dd/mm/yyyy"" 00"":""00"":""00"""
  ) AS updatedate,
  Format(
    [Date], "dd/mm/yyyy"" 00"":""00"":""00"""
  ) AS invoicedate,
  Tbl_Invoices.Document_Number AS Expr2,
  Tbl_Invoices.Customer_reference AS Expr3,
  Tbl_Invoices.SONumber AS Expr4,
  Tbl_Invoices.Type AS Expr5,
  Tbl_Invoices.OriginalAmount AS Expr6,
  Tbl_Invoices.Amount AS Expr7,
  Format(
    [Overdue_Date], "dd/mm/yyyy"" 00"":""00"":""00"""
  ) AS invoiceduedate,
  Tbl_Invoices.Currency AS Expr8
FROM
  Tbl_Invoices
WHERE
  (
    (
      ([Tbl_Invoices].[Update_date])= Date()
    )
  );
