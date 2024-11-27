SELECT
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Users.Name,
  Tbl_Customers.Country
FROM
  Tbl_Users
  INNER JOIN Tbl_Customers ON Tbl_Users.ID = Tbl_Customers.Credit_controller
ORDER BY
  Tbl_Users.Name;
