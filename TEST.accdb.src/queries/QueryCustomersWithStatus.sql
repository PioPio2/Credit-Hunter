SELECT
  Tbl_Users.Name,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name AS [Customer name],
  Tbl_Customers.Status,
  Tbl_Customers.StatusDate
FROM
  Tbl_Customers
  LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID
WHERE
  (
    (
      (
        Tbl_Customers.Credit_controller
      )= GetNumCreditController(
        fOSUserName()
      )
    )
  )
GROUP BY
  Tbl_Users.Name,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Customers.Status,
  Tbl_Customers.StatusDate
HAVING
  (
    (
      (Tbl_Customers.Status) Is Not Null
    )
  )
ORDER BY
  Tbl_Customers.Name;
