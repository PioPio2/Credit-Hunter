SELECT
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Country,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Date,
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Type,
  Tbl_Invoices.Amount,
  Tbl_Invoices.Currency,
  Tbl_Invoices.Query,
  Tbl_Invoices.mEMO
FROM
  Tbl_Customers
  INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (
        Tbl_Customers.Credit_controller
      )= GetNumCreditController(
        fOSUserName()
      )
    )
    AND (
      (Tbl_Invoices.Type)= 4
      Or (Tbl_Invoices.Type)= 2
    )
    AND (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
ORDER BY
  Tbl_Customers.Country,
  Tbl_Customers.Name;
