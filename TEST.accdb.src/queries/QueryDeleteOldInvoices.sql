DELETE Tbl_Invoices.Update_date
FROM
  Tbl_Invoices
WHERE
  (
    (
      (Tbl_Invoices.Update_date)< DateAdd(
        "d",
        -31,
        Now()
      )
    )
  );
