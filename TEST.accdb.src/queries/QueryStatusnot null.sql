SELECT
  Tbl_Customers.Customer_code,
  Tbl_Customer_Status.Status,
  Tbl_Customers.Name
FROM
  Tbl_Customers
  INNER JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID
WHERE
  (
    (
      (Tbl_Customer_Status.Status) Is Not Null
    )
    AND (
      (
        Tbl_Customers.Credit_controller
      )= GetNumCreditController(
        fOSUserName()
      )
    )
  )
ORDER BY
  Tbl_Customer_Status.Status,
  Tbl_Customers.Name;
