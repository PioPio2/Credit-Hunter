VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWhoPaidYesterdaySupportLIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IWhoPaidYesterdaySupport

Private Sub IWhoPaidYesterdaySupport_SetLastInvoiceImportDate(aDate As Date)

End Sub

Private Sub IWhoPaidYesterdaySupport_SetOneBeforeLastInvoiceImportDate(aDate As Date)

End Sub

Public Function IWhoPaidYesterdaySupport_GetLastInvoiceImportDate(SQL As String) As Date
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQL)
    IWhoPaidYesterdaySupport_GetLastInvoiceImportDate = RS.Fields(0).value
End Function

Public Function IWhoPaidYesterdaySupport_OneBeforeLastInvoiceImportDate(SQL As String) As Date
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQL)
    IWhoPaidYesterdaySupport_OneBeforeLastInvoiceImportDate = RS.Fields(0).value
End Function

Public Function IWhoPaidYesterdaySupport_GetWhoPaidYesterdayRecordset(SQL As String) As Long
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQL)
End Function

Public Function IWhoPaidYesterdaySupport_UpdateQueryInvoicesLastDate(SQL As String) As Boolean
    CurrentDb.QueryDefs("QueryInvoicesLastDate").SQL = SQL
End Function

Public Function IWhoPaidYesterdaySupport_UpdateQueryInvoicesPreviousDate(SQL As String) As Boolean
    CurrentDb.QueryDefs("QueryInvoicesPreviousDate").SQL = SQL
End Function
