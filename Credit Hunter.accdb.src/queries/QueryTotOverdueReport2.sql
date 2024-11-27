SELECT
  Tbl_Historical_Statements.Update_date AS Expr1,
  Tbl_Historical_Statements.Customer_ID AS Expr2,
  Sum(
    Tbl_Historical_Statements.Amount
  ) AS SumOfAmount
FROM
  Tbl_Historical_Statements
WHERE
  (
    (
      (
        [Tbl_Historical_Statements].[Overdue_Date]
      )<= DateDiff("d", 30, [Update_date])
    )
  )
GROUP BY
  Tbl_Historical_Statements.Update_date,
  Tbl_Historical_Statements.Customer_ID;
