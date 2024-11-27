SELECT
  QueryChargebacksOpenNow.*,
  QueryChargebacksOpenPreviousDate.Customer_ID,
  Tbl_Customers.Name,
  QueryChargebacksOpenNow.Document_Number
FROM
  Tbl_Customers
  INNER JOIN (
    QueryChargebacksOpenPreviousDate
    RIGHT JOIN QueryChargebacksOpenNow ON (
      QueryChargebacksOpenPreviousDate.Document_Number = QueryChargebacksOpenNow.Document_Number
    )
    AND (
      QueryChargebacksOpenPreviousDate.Customer_ID = QueryChargebacksOpenNow.Customer_ID
    )
  ) ON Tbl_Customers.Customer_code = QueryChargebacksOpenNow.Customer_ID
WHERE
  (
    (
      (
        QueryChargebacksOpenPreviousDate.Customer_ID
      ) Is Null
    )
  )
ORDER BY
  QueryChargebacksOpenNow.Document_Number;
