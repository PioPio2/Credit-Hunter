DELETE Query6.Update_date,
Tbl_InvoiceAttachments.*
FROM
  Tbl_InvoiceAttachments
  LEFT JOIN Query6 ON Tbl_InvoiceAttachments.DocumentID = Query6.a
WHERE
  (
    (
      (Query6.Update_date) Is Null
    )
  );
