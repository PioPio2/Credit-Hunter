SELECT
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Date,
  Tbl_Invoices.Currency,
  Sum(Tbl_Invoices.Amount) AS SumOfAmount
FROM
  Tbl_Customers
  INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (Tbl_Invoices.Type)= 3
    )
    AND (
      (Tbl_Invoices.Update_date)=(
        SELECT
          Max(Tbl_Invoices.Update_date) AS MaxOfUpdate_date
        FROM
          Tbl_Invoices;
      )
    )
  )
GROUP BY
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Date,
  Tbl_Invoices.Currency
HAVING
  (
    (
      (
        Tbl_Customers.Credit_controller
      )= GetNumCreditController(
        fOSUserName()
      )
    )
  )
ORDER BY
  Sum(Tbl_Invoices.Amount);
