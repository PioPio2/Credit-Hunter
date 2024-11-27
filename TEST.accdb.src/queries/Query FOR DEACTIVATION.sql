SELECT
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_Customers.Country,
  Tbl_Areas.Area,
  Tbl_Users.Name
FROM
  (
    Tbl_Customers
    LEFT JOIN Tbl_Areas ON Tbl_Customers.Area = Tbl_Areas.ID
  )
  LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID;
