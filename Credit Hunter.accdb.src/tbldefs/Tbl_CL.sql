CREATE TABLE [Tbl_CL] (
  [Customer_code] LONG CONSTRAINT [Customer_code] UNIQUE,
  [Currency] VARCHAR (3),
  [CreditLimit] CURRENCY,
  [OpenARBalance] CURRENCY,
  [AwaitingInvoicing] CURRENCY,
  [AmtScheduledTom] CURRENCY,
  [AmtScheduled8Dyas] CURRENCY
)
