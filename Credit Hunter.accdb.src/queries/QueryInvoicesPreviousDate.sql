SELECT
  Tbl_Invoices.Update_date,
  Tbl_Invoices.*,
  Tbl_Types.Descripition
FROM
  Tbl_Invoices
  LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID
WHERE
  (
    (
      (Tbl_Invoices.Update_date)= #10/12/2024#
    )
  );
