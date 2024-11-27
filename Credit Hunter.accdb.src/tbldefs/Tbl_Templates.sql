CREATE TABLE [Tbl_Templates] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [Language] LONG,
  [Text] LONGTEXT,
  [Step] LONG,
  [TemplateName] VARCHAR (50),
  [Subject] VARCHAR (255)
)
