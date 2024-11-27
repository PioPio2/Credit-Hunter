SELECT
  Tbl_Users.Name,
  Tbl_Customers.Name AS [Customer name],
  Tbl_Customer_Status.Status,
  Tbl_Customers.StatusDate,
  DateDiff(
    "d",
    [StatusDate],
    Date()
  ) AS [Days gone by],
  [QueryCustomerOverdue>30daysOnlytotals].Expr1,
  Tbl_Customers.Customer_code
FROM
  (
    (
      Tbl_Customers
      LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID
    )
    INNER JOIN [QueryCustomerOverdue>30daysOnlytotals] ON Tbl_Customers.Customer_code = [QueryCustomerOverdue>30daysOnlytotals].Customer_code
  )
  LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID
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
      (
        [QueryCustomerOverdue>30daysOnlytotals].Expr1
      )> 0
    )
  )
GROUP BY
  Tbl_Users.Name,
  Tbl_Customers.Name,
  Tbl_Customer_Status.Status,
  Tbl_Customers.StatusDate,
  [QueryCustomerOverdue>30daysOnlytotals].Expr1,
  Tbl_Customers.Customer_code
ORDER BY
  [QueryCustomerOverdue>30daysOnlytotals].Expr1 DESC;
