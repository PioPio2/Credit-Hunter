SELECT
  Tbl_Customers.Credit_controller,
  Tbl_Customers.RetailOEM
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
  Tbl_Customers.Credit_controller,
  Tbl_Customers.RetailOEM;
