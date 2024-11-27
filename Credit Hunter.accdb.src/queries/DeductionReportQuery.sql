SELECT
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Type
FROM
  Tbl_Customers
  INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
GROUP BY
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Type
HAVING
  (
    (
      (Tbl_Invoices.Type)= 4
    )
  )
ORDER BY
  Tbl_Customers.Name;
