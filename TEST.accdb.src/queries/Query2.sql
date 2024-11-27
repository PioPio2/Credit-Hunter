SELECT
  Tbl_Customers.Customer_code,
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Name,
  Query1.[AR Exposure in main currency],
  Query1.Update_date,
  Query1.[Overdue 90+ days],
  Query1.ExchangeRate,
  Tbl_Customer_Status.Status,
  Query1.[Total overdue on fiscal month end],
  Tbl_Customers.Country,
  Query1.[Overdue as of Today in main currency],
  Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Target in Main currency],
  IIf(
    IsNull([AmountInEUR]),
    0,
    [AmountInEUR]
  ) AS [Already collected in Main currency],
  [MonthlyTargetInMainCurrency] - [Already collected in Main currency] AS [Still to be collected in Main currency],
  IIf(
    [MonthlyTargetInMainCurrency] <> 0,
    IIf(
      IsNull([AmountInEUR]),
      0,
      [AmountInEUR]
    )* 100 / [MonthlyTargetInMainCurrency],
    0
  ) AS [% Cash Target Achieved]
FROM
  (
    (
      Tbl_Customers
      LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID
    )
    LEFT JOIN QueryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyCollectedInEUR.CustomerID
  )
  LEFT JOIN Query1 ON Tbl_Customers.Customer_code = Query1.Customer_ID
GROUP BY
  Tbl_Customers.Customer_code,
  Tbl_Customers.Credit_controller,
  Tbl_Customers.Name,
  Query1.[AR Exposure in main currency],
  Query1.Update_date,
  Query1.[Overdue 90+ days],
  Query1.ExchangeRate,
  Tbl_Customer_Status.Status,
  Query1.[Total overdue on fiscal month end],
  Tbl_Customers.Country,
  Query1.[Overdue as of Today in main currency],
  Tbl_Customers.MonthlyTargetInMainCurrency,
  IIf(
    IsNull([AmountInEUR]),
    0,
    [AmountInEUR]
  ),
  IIf(
    [MonthlyTargetInMainCurrency] <> 0,
    IIf(
      IsNull([AmountInEUR]),
      0,
      [AmountInEUR]
    )* 100 / [MonthlyTargetInMainCurrency],
    0
  )
HAVING
  (
    (
      (
        Tbl_Customers.Credit_controller
      )= 1
    )
  );
