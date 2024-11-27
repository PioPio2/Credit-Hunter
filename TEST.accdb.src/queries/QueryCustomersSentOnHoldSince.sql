SELECT
  Tbl_Customers.Name,
  Tbl_Customers.Customer_code,
  TblNotes.CustomerCode,
  TblNotes.Note
FROM
  Tbl_Customers
  LEFT JOIN TblNotes ON Tbl_Customers.Customer_code = TblNotes.CustomerCode
WHERE
  (
    (
      (TblNotes.ID)> 55411
    )
  )
GROUP BY
  Tbl_Customers.Name,
  Tbl_Customers.Customer_code,
  TblNotes.CustomerCode,
  TblNotes.Note
HAVING
  (
    (
      (TblNotes.Note) Like "*Modify status into: HOLD*"
    )
  );
