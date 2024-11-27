SELECT
  Tbl_Customers.Customer_code,
  Tbl_Invoices.Currency,
  Sum(Tbl_Invoices.Amount) AS ARExposure,
  Sum(
    IIf(
      [Overdue_Date] < Date(),
      [amount],
      0
    )
  ) AS AROverdue,
  Tbl_CL.CreditLimit AS [CreditLimit in EUR]
FROM
  Tbl_Customers
  LEFT JOIN (
    Tbl_Invoices
    LEFT JOIN Tbl_CL ON Tbl_Invoices.Customer_ID = Tbl_CL.Customer_code
  ) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
GROUP BY
  Tbl_Customers.Customer_code,
  Tbl_Invoices.Currency,
  Tbl_CL.CreditLimit;
