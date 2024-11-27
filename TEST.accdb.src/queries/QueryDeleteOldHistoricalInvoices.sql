DELETE Tbl_Invoices_History.Date AS Expr1,
[Tbl_Invoices_History].[Date] AS Expr2
FROM
  Tbl_Invoices_History
WHERE
  (
    (
      ([Tbl_Invoices_History].[Date])< DateAdd(
        "yyyy",
        -1,
        Now()
      )
    )
  );
