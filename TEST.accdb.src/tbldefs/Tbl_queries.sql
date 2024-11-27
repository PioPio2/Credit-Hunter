CREATE TABLE [Tbl_queries] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [Query] VARCHAR (50),
  [Resolution_owner] VARCHAR (50),
  [InvoiceToBePaid] BIT,
  [ToFillChargebackFile] BIT
)
