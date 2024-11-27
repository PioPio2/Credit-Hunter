VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWhoPaidYesterday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IWhoPaidYesterday

Private Support As IWhoPaidYesterdaySupport

Private Function IWhoPaidYesterday_GetRecordSourceSQL() As String
    Dim SQL As String
    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "        DISTINCTROW Tbl_Customers.* "
    SQL = SQL & "FROM "
    SQL = SQL & "        Tbl_Customers "
    SQL = SQL & "INNER JOIN "
    SQL = SQL & "        QueryInvoicesClosedInLastDate "
    SQL = SQL & "ON "
    SQL = SQL & "        Tbl_Customers.Customer_code = QueryInvoicesClosedInLastDate.Customer_ID "
    SQL = SQL & "WHERE "
    SQL = SQL & "        ( "
    SQL = SQL & "                ( "
    SQL = SQL & "                        ( "
    SQL = SQL & "                                Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ")) "
    'SQL = SQL & "ORDER BY QueryInvoicesLastDate.Date;"
    IWhoPaidYesterday_GetRecordSourceSQL = SQL
End Function

Public Function IWhoPaidYesterday_Inject(ByRef aInject As IWhoPaidYesterdaySupport) As Boolean
    Set Support = aInject
    IWhoPaidYesterday_Inject = True
End Function

Private Function IWhoPaidYesterday_LastInvoiceImportDateSQL() As String
    Dim SQL As String
    '    Dim RS As DAO.Recordset
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        Max([Tbl_Invoices].[Update_date]) As [Max_Update_date] "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Invoices]"
    IWhoPaidYesterday_LastInvoiceImportDateSQL = SQL
End Function

Private Function IWhoPaidYesterday_QueryLastInvoiceImportSQL() As String
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        Max([Tbl_Invoices].[Update_date]) As Max_Update_date "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Invoices] "
    SQL = SQL & "Where "
    SQL = SQL & "        [Tbl_Invoices].[Update_date] < "
    SQL = SQL & "        ( "
    SQL = SQL & "                Select "
    SQL = SQL & "                        Max([Tbl_Invoices].[Update_date]) As Max_Update_date "
    SQL = SQL & "                From "
    SQL = SQL & "                        [Tbl_Invoices])"

    IWhoPaidYesterday_QueryLastInvoiceImportSQL = SQL
End Function

Private Function IWhoPaidYesterday_BeforeLastInvoiceImportDateSQL() As String
    Dim SQL As String
    SQL = IWhoPaidYesterday_QueryLastInvoiceImportSQL ' find SQL for the one before last date the data was imported in tbl_invoices

    Dim OneBeforeLastInvoiceImportDate As Date
    OneBeforeLastInvoiceImportDate = Support.OneBeforeLastInvoiceImportDate(SQL) ' find out the date of the one before last import

    Dim QueryDate As String
    QueryDate = Format$(OneBeforeLastInvoiceImportDate, "\#mm\/dd\/yyyy\#") ' transform the data in US format (Access works with US format only for queries

    IWhoPaidYesterday_BeforeLastInvoiceImportDateSQL = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=" & QueryDate & "));"
End Function

Public Function IWhoPaidYesterday_QueryInvoicesLastDateSQL() As String
    Dim SQL As String
    SQL = IWhoPaidYesterday_LastInvoiceImportDateSQL ' find SQL for the last date the data was imported in tbl_invoices

    Dim LastInvoiceImportDate As Date
    LastInvoiceImportDate = Support.GetLastInvoiceImportDate(SQL) ' find the date of the last import in tbl_invoices

    Dim QueryDate As String
    QueryDate = Format$(LastInvoiceImportDate, "\#mm\/dd\/yyyy\#") ' transform the data in US format (Access works with US format only for queries

    IWhoPaidYesterday_QueryInvoicesLastDateSQL = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=" & QueryDate & "));"
End Function


Private Function IWhoPaidYesterday_UpdateSecondaryQueries() As Boolean
    Dim SQLString As String
    SQLString = IWhoPaidYesterday_QueryInvoicesLastDateSQL
    Support.UpdateQueryInvoicesLastDate (SQLString)
    'CurrentDb.QueryDefs("QueryInvoicesLastDate").SQL = SQLString

    SQLString = IWhoPaidYesterday_BeforeLastInvoiceImportDateSQL
    Support.UpdateQueryInvoicesPreviousDate (SQLString)
    'CurrentDb.QueryDefs("QueryInvoicesPreviousDate").SQL = SQLString

    IWhoPaidYesterday_UpdateSecondaryQueries = True
End Function
