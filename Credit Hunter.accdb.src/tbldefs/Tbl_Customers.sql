﻿CREATE TABLE [Tbl_Customers] (
  [Customer_code] VARCHAR (255) CONSTRAINT [Customer_code] UNIQUE,
  [Credit_controller] LONG,
  [Name] VARCHAR (50),
  [Address] VARCHAR (50),
  [Address2] VARCHAR (50),
  [Address3] VARCHAR (50),
  [Address4] VARCHAR (50),
  [Address5] VARCHAR (50),
  [Country] VARCHAR (2),
  [Update_date] DATETIME,
  [OWN_company] VARCHAR (50),
  [OWN_bank_details1] VARCHAR (50),
  [OWN_bank_details2] VARCHAR (50),
  [OWN_bank_details3] VARCHAR (50),
  [OWN_bank_details4] VARCHAR (50),
  [NextAppointment] DATETIME,
  [Email] VARCHAR (150),
  [DA TOGLIEREEEEEEEE TextEmail] LONGTEXT,
  [ccEmail] VARCHAR (255),
  [StatusDate] DATETIME,
  [ToSendStatement] BIT,
  [Index] CURRENCY,
  [Note] LONGTEXT,
  [ToSendRequestRelease] BIT,
  [TotalInsurance] CURRENCY,
  [StatementForm] UNSIGNED BYTE,
  [RetailOEM] VARCHAR (10),
  [ToReleaseOrder] BIT,
  [Status] LONG,
  [Timezone] LONG,
  [Language] LONG,
  [ContactNames] VARCHAR (150),
  [DSO] LONG,
  [EmailCode] LONG,
  [Area] LONG,
  [LastStatementSent] DATETIME,
  [FacturaNumberToBePrinted] BIT,
  [PullTicketNumberToBePrinted] BIT,
  [OriginalInvoiceAmountToBePrinted] BIT,
  [HighestExposure] CURRENCY,
  [ReleaseNotes] LONGTEXT,
  [MonthlyTargetInMainCurrency] CURRENCY,
  [Credit Limit] CURRENCY,
  [MainPhoneNumber] VARCHAR (255)
)