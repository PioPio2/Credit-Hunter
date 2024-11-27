SELECT
  Tbl_Users.Name,
  Tbl_Customers.RetailOEM,
  Tbl_Customers.Name,
  Tbl_CashCollected.[Payment Date],
  Sum(Tbl_CashCollected.Amount) AS [Amount in EUR],
  Tbl_CashCollected.Currency,
  Tbl_CashCollected.[Original amount],
  Tbl_Users.ID
FROM
  (
    (
      Tbl_Customers
      INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID
    )
    INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID
  )
  INNER JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID
GROUP BY
  Tbl_Users.Name,
  Tbl_Customers.RetailOEM,
  Tbl_Customers.Name,
  Tbl_CashCollected.[Payment Date],
  Tbl_CashCollected.Currency,
  Tbl_CashCollected.[Original amount],
  Tbl_Users.ID
HAVING
  (
    (
      (
        Tbl_CashCollected.[Payment Date]
      )>= #02/21/15#
      And (
        Tbl_CashCollected.[Payment Date]
      )<= #03/31/15#
    )
  )
ORDER BY
  Sum(Tbl_CashCollected.Amount) DESC;
