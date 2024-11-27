VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private TotalOutstanding As Currency
Private TotalOverdue As Currency
Private aCurrency As String

Function Populate(aTotalOutstanding, aTotalOverdue As Currency, inCurrency As String) As Boolean
    TotalOutstanding = aTotalOutstanding
    TotalOverdue = aTotalOverdue
    aCurrency = inCurrency
    Populate = True
End Function

Function getTotalOutstanding() As Currency
    getTotalOutstanding = TotalOutstanding
End Function

Function getTotalOverdue() As Currency
    getTotalOverdue = TotalOverdue
End Function

Function getCurrency() As String
    getCurrency = aCurrency
End Function
