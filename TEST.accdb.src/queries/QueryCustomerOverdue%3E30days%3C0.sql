SELECT
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Sum(
    [QueryCustomerOverdue>30daysOnlytotals].Expr1
  ) AS aa,
  Tbl_Customer_Status.Status
FROM
  (
    Tbl_Customers
    INNER JOIN [QueryCustomerOverdue>30daysOnlytotals] ON Tbl_Customers.Customer_code = [QueryCustomerOverdue>30daysOnlytotals].Customer_code
  )
  INNER JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID
GROUP BY
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Customer_Status.Status
HAVING
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
        Sum(
          [QueryCustomerOverdue>30daysOnlytotals].Expr1
        )
      )<= 0
      Or (
        Sum(
          [QueryCustomerOverdue>30daysOnlytotals].Expr1
        )
      ) Is Null
    )
    AND (
      (Tbl_Customer_Status.Status) Is Not Null
    )
  )
ORDER BY
  Tbl_Customer_Status.Status;
