CREATE TABLE [Tbl_Invoices_History] (
  [Customer_ID] LONG,
  [Update_date] DATETIME,
  [Date] DATETIME,
  [Document_Number] VARCHAR (12),
  [Customer_reference] VARCHAR (50),
  [Type] LONG,
  [Amount] CURRENCY,
  [Overdue_Date] DATETIME,
  [Currency] VARCHAR (3),
  [PaymentDate] DATETIME,
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE
)
