CREATE TABLE [Tbl_Historical_Statements] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [Customer_ID] LONG,
  [Update_date] DATETIME,
  [Date] DATETIME,
  [Document_Number] VARCHAR (12),
  [Customer_reference] VARCHAR (50),
  [Type] LONG,
  [Amount] CURRENCY,
  [Overdue_Date] DATETIME,
  [Currency] VARCHAR (3),
  [Query] LONG,
  [mEMO] LONGTEXT
)
