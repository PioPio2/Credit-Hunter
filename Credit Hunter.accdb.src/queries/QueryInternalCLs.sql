SELECT
  Tbl_Countries.Country,
  Tbl_Countries.Area,
  Tbl_Customers.Country,
  Tbl_Customers.Customer_code,
  Tbl_Customers.Name,
  Tbl_CL.CreditLimit,
  Tbl_Customers.Status
FROM
  (
    Tbl_Customers
    LEFT JOIN Tbl_CL ON Tbl_Customers.Customer_code = Tbl_CL.Customer_code
  )
  LEFT JOIN Tbl_Countries ON Tbl_Customers.Country = Tbl_Countries.Code
ORDER BY
  Tbl_Countries.Area;
