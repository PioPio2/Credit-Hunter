CREATE TABLE [Tbl_PaymentData] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [CustomerCode] LONG,
  [PaymentIncoming] CURRENCY,
  [PaymentDate] DATETIME,
  [Proof] VARCHAR (50),
  [Comment] LONGTEXT
)
