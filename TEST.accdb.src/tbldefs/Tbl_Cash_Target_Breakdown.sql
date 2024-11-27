CREATE TABLE [Tbl_Cash_Target_Breakdown] (
  [CControllerID] LONG,
  [Channel] VARCHAR (25),
  [OriginalCurrency] VARCHAR (5),
  [FiscalYear] LONG,
  [FiscalQuarter] LONG,
  [FiscalMonth] LONG,
  [Amount] CURRENCY,
  [ChannelCurrency] VARCHAR (5),
  [ExchangeRateToMainCurrency] SINGLE,
  [AmountInUSD] CURRENCY
)
