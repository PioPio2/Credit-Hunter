CREATE TABLE [Tbl_CashCollected] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [CustomerID] VARCHAR (255),
  [Payment Date] DATETIME,
  [Currency] VARCHAR (3),
  [Amount] CURRENCY,
  [Original amount] CURRENCY,
  [RETnumber] VARCHAR (10),
  [FiscalYear] LONG,
  [FiscalMonth] LONG,
  [FiscalQuarter] LONG,
  [PaymentStillAvailable] BIT,
  [AmountInUSD] CURRENCY,
  [InvoiceNumber] VARCHAR (255)
)
