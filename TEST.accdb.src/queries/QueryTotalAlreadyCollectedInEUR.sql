SELECT
  Sum(
    (
      [ExchangeRate] * [Original amount]
    )
  ) AS AmountInEUR,
  Tbl_CashCollected.CustomerID
FROM
  Tbl_CashCollected
  INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID
WHERE
  (
    (
      (
        Tbl_CashCollected.[Payment Date]
      )>= DateAdd(
        "d",
        1,
        DMax(
          "[MonthEnd]",
          "[Tbl_MonthEnd]",
          "MonthEnd <#" & Format(
            Date(),
            "mm/dd/yy"
          )& "#"
        )
      )
    )
  )
GROUP BY
  Tbl_CashCollected.CustomerID;
