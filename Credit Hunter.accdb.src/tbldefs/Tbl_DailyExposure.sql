CREATE TABLE [Tbl_DailyExposure] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [CustomerID] LONG,
  [RefDate] DATETIME,
  [Currency] VARCHAR (3),
  [ARExposure] CURRENCY,
  [TotalOverdue] CURRENCY,
  [OracleCL] CURRENCY
)
