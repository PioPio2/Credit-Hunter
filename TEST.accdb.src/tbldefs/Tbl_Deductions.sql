CREATE TABLE [Tbl_Deductions] (
  [ID] AUTOINCREMENT,
  [CustomerID] LONG,
  [DeductionDate] DATETIME,
  [Currency] VARCHAR (3),
  [Amount] CURRENCY,
  [$martn#] VARCHAR (50),
  [Chargeback n#] LONGTEXT,
  [OffsetAgainstCM] CURRENCY,
  [Note] LONGTEXT
)
