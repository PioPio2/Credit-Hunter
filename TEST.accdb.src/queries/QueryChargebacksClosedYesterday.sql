SELECT
  QueryChargebacksOpenNow.Customer_ID,
  Tbl_Customers.Name,
  QueryChargebacksOpenPreviousDate.Document_Number,
  QueryChargebacksOpenPreviousDate.*
FROM
  (
    QueryChargebacksOpenPreviousDate
    LEFT JOIN QueryChargebacksOpenNow ON (
      QueryChargebacksOpenPreviousDate.Document_Number = QueryChargebacksOpenNow.Document_Number
    )
    AND (
      QueryChargebacksOpenPreviousDate.Customer_ID = QueryChargebacksOpenNow.Customer_ID
    )
  )
  INNER JOIN Tbl_Customers ON QueryChargebacksOpenPreviousDate.Customer_ID = Tbl_Customers.Customer_code
WHERE
  (
    (
      (
        QueryChargebacksOpenNow.Customer_ID
      ) Is Null
    )
  )
ORDER BY
  QueryChargebacksOpenPreviousDate.Document_Number;
