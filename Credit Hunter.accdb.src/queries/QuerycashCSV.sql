SELECT
  Tbl_CashCollected.CustomerID AS Expr1,
  Format(
    [Payment Date], "dd/mm/yyyy"" 00"":""00"":""00"""
  ) AS a,
  Tbl_CashCollected.Currency AS Expr2,
  Tbl_CashCollected.Amount AS Expr3,
  Tbl_CashCollected.[Original amount] AS Expr4,
  Tbl_CashCollected.RETnumber AS Expr5,
  Tbl_CashCollected.FiscalYear AS Expr6,
  Tbl_CashCollected.FiscalMonth AS Expr7,
  Tbl_CashCollected.FiscalQuarter AS Expr8
FROM
  Tbl_CashCollected;
