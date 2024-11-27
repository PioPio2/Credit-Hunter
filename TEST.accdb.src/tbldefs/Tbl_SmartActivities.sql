CREATE TABLE [Tbl_SmartActivities] (
  [CustomerID] LONG,
  [ClaimID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [$martn#] VARCHAR (35),
  [CMn] VARCHAR (12),
  [Currency] VARCHAR (3),
  [Amount] CURRENCY,
  [ApprovedAmount] CURRENCY,
  [AlreadyOffsetAgainstDeduction] DATETIME,
  [OffsetDate] DATETIME,
  [s_ColLineage] LONGBINARY,
  [s_Generation] LONG,
  [s_GUID] GUID CONSTRAINT [s_GUID] UNIQUE,
  [s_Lineage] LONGBINARY
)
