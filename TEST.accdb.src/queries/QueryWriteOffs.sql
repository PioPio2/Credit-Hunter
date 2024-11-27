SELECT
  Tbl_Users.Name,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Date AS [Document date],
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Currency,
  Tbl_Invoices.Amount,
  Tbl_Invoices.Overdue_Date AS [Due date],
  Tbl_Customer_Status.Status,
  Tbl_Invoices.Query,
  Tbl_Invoices.mEMO AS [Note]
FROM
  (
    (
      Tbl_Customers
      INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
    )
    LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID
  )
  LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= Date()
      Or (Tbl_Invoices.Update_date)= Date()
    )
  )
GROUP BY
  Tbl_Users.Name,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Invoices.Date,
  Tbl_Invoices.Document_Number,
  Tbl_Invoices.Currency,
  Tbl_Invoices.Amount,
  Tbl_Invoices.Overdue_Date,
  Tbl_Customer_Status.Status,
  Tbl_Invoices.Query,
  Tbl_Invoices.mEMO
HAVING
  (
    (
      (Tbl_Invoices.Date)<=(
        DMin(
          "[MonthEnd]",
          "[Tbl_MonthEnd]",
          "[MonthEnd]>=#" & Format(
            Date(),
            "mm/dd/yy"
          )& "#"
        )
      )-180
    )
  )
  OR (
    (
      (Tbl_Invoices.Date)>(
        DMin(
          "[MonthEnd]",
          "[Tbl_MonthEnd]",
          "[MonthEnd]>=#" & Format(
            Date(),
            "mm/dd/yy"
          )& "#"
        )
      )-180
    )
    AND (
      (Tbl_Invoices.Query)= 19
    )
  )
ORDER BY
  Tbl_Users.Name,
  Tbl_Customers.Name;
