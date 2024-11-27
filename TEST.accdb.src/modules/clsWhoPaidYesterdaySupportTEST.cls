VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWhoPaidYesterdaySupportTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IWhoPaidYesterdaySupport
Private aLastInvoiceImportDate As Date
Private aOneBeforeLastInvoiceImportDate As Date

Public Function IWhoPaidYesterdaySupport_GetLastInvoiceImportDate(SQL As String) As Date
    IWhoPaidYesterdaySupport_GetLastInvoiceImportDate = aLastInvoiceImportDate
End Function

Private Function IWhoPaidYesterdaySupport_OneBeforeLastInvoiceImportDate(SQL As String) As Date
    IWhoPaidYesterdaySupport_OneBeforeLastInvoiceImportDate = aOneBeforeLastInvoiceImportDate
End Function

Private Sub IWhoPaidYesterdaySupport_SetLastInvoiceImportDate(aDate As Date)
    aLastInvoiceImportDate = aDate
End Sub

Private Sub IWhoPaidYesterdaySupport_SetOneBeforeLastInvoiceImportDate(aDate As Date)
    aOneBeforeLastInvoiceImportDate = aDate
End Sub


Private Function IWhoPaidYesterdaySupport_UpdateQueryInvoicesLastDate(SQL As String) As Boolean

End Function

Private Function IWhoPaidYesterdaySupport_UpdateQueryInvoicesPreviousDate(SQL As String) As Boolean

End Function
