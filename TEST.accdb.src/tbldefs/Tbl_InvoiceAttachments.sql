CREATE TABLE [Tbl_InvoiceAttachments] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [AttachName] VARCHAR (255),
  [CustomerID] LONG,
  [DocumentID] VARCHAR (60)
)
