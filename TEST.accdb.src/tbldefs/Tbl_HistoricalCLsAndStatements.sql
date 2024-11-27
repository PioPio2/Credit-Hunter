CREATE TABLE [Tbl_HistoricalCLsAndStatements] (
  [Update_date] DATETIME,
  [Customer_code] LONG,
  [Currency] VARCHAR (3),
  [CreditLimit] CURRENCY,
  [OpenARBalance] CURRENCY,
  [AwaitingInvoicing] CURRENCY,
  [AmtScheduledTom] CURRENCY,
  [AmtScheduled8Dyas] CURRENCY,
  [Current] CURRENCY,
  [Overdue1-30Days] CURRENCY,
  [Overdue31-60Days] CURRENCY,
  [Overdue61-90Days] CURRENCY,
  [OverdueOver90Days] CURRENCY,
  [InsuranceCreditLimit] CURRENCY
)
