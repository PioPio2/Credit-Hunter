﻿SELECT
  Tbl_Customers.Name,
  Tbl_Invoices.Date,
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Amount,
  Tbl_Invoices.Overdue_Date,
  Tbl_Invoices.mEMO
FROM
  Tbl_Customers
  INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (Tbl_Customers.Name) Like '*a*'
      OR (Tbl_Customers.Name) Like '*c*'
      OR (Tbl_Customers.Name) Like '*b*'
    )
    AND (
      (Tbl_Invoices.Update_date)= Date()
    )
  )
ORDER BY
  Tbl_Customers.Name,
  Tbl_Invoices.Date;