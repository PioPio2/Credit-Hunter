CREATE TABLE [Tbl_Users] (
  [ID] AUTOINCREMENT CONSTRAINT [ID] UNIQUE,
  [UserName] VARCHAR (15),
  [Name] VARCHAR (50),
  [EmailText] LONGTEXT,
  [Querywithoutcreditcontroller] BIT,
  [Onaccountsstillopen] BIT,
  [Whopaidyesterdayroutine] BIT,
  [QuerywithoutcreditcontrollerEvery] VARCHAR (50),
  [OnaccountsstillopenEvery] VARCHAR (50),
  [WhopaidyesterdayroutineEvery] VARCHAR (50),
  [Signature] LONGTEXT,
  [Superuser] BIT,
  [E-mailAddress] VARCHAR (255),
  [Password] VARCHAR (255),
  [RetypePassword] VARCHAR (255),
  [EmailInterface] VARCHAR (255),
  [EmailSentToSender] BIT
)
