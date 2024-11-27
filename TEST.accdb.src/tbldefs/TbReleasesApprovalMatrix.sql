CREATE TABLE [TbReleasesApprovalMatrix] (
  [LevelNumber] AUTOINCREMENT CONSTRAINT [LevelNumber] UNIQUE,
  [ApprovalLimit] CURRENCY,
  [EmailAddress] VARCHAR (90),
  [Name] VARCHAR (50)
)
