SELECT
  Tbl_Invoices.*,
  Tbl_Types.ToFillChargbackFile,
  Tbl_Invoices.Update_date
FROM
  Tbl_Invoices
  INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID
WHERE
  (
    (
      (Tbl_Types.ToFillChargbackFile)= True
    )
    And (
      (Tbl_Invoices.Update_date)= #4/6/2010#
    )
  );
