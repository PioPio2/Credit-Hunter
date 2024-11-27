CREATE TABLE [Tbl_Customer_Status] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [Description] VARCHAR (50),
  [Step] LONG,
  [Status] VARCHAR (50),
  [AppearsInTheScheduler] BIT,
  [ToSendStatement] BIT,
  [ToSendEmail] BIT
)
