SELECT
  Tbl_Customers.Credit_controller AS Expr1,
  Tbl_Customers.Customer_code AS Expr2,
  Tbl_Customers.Name AS Expr3,
  Tbl_Customers.Country AS Expr4
FROM
  Tbl_Customers
WHERE
  (
    (
      (
        [Tbl_Customers].[Credit_controller]
      ) Is Null
    )
  )
ORDER BY
  Tbl_Customers.Name;
