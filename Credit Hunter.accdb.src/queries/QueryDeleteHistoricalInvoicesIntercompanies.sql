DELETE Tbl_Invoices_History.*,
Tbl_Users.UserName
FROM
  Tbl_Users
  INNER JOIN (
    Tbl_Customers
    INNER JOIN Tbl_Invoices_History ON Tbl_Customers.Customer_code = Tbl_Invoices_History.Customer_ID
  ) ON Tbl_Users.ID = Tbl_Customers.Credit_controller
WHERE
  (
    (
      (Tbl_Users.UserName)= "int"
    )
  );
