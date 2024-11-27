SELECT
  Tbl_Invoices.Customer_ID AS Expr1,
  Tbl_Invoices.Update_date AS Expr2,
  Tbl_Invoices.Date AS Expr3,
  Tbl_Invoices.Document_Number AS Expr4,
  Tbl_Invoices.QueryToBePrinted AS Expr5
FROM
  Tbl_Invoices
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= #3/31/2015#
    )
    And (
      (Tbl_Invoices.QueryToBePrinted)= No
    )
  );
