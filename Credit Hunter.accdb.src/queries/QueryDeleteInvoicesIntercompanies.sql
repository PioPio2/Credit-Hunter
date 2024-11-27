DELETE Tbl_Users.UserName,
Tbl_Invoices.*
FROM
  (
    Tbl_Users
    INNER JOIN Tbl_Customers ON Tbl_Users.ID = Tbl_Customers.Credit_controller
  )
  INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID
WHERE
  (
    (
      (Tbl_Users.UserName)= "int"
    )
  );
