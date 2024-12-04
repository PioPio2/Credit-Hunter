VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatementSupportLIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements iStatementSupport

Private ExcelTemplate As Excel.Workbook
Private Function iStatementSupport_CloseStatement() As Boolean
    ExcelTemplate.Close
    iStatementSupport_CloseStatement = True
End Function

Private Function iStatementSupport_CopyStatementData(aCellStart As Range, aData As ADODB.Recordset, ByRef outOutstanding As Currency, ByRef outOverdue As Currency, ByVal NRec As Long) As Long
    Dim DummyRng As String
    DummyRng = aCellStart.Address
    ExcelTemplate.Sheets(1).Range(DummyRng).value = ""
    Dim I As Long
    For I = 1 To NRec
        ExcelTemplate.Sheets(1).Range(DummyRng).EntireRow.Insert
        ExcelTemplate.Sheets(1).Range(DummyRng).EntireRow.Font.Bold = False
    Next I

    ExcelTemplate.Sheets(1).Range(DummyRng).CopyFromRecordset aData
    Dim Region As Variant

    Dim Rng As Range
    Set Rng = ExcelTemplate.Sheets(1).Range(DummyRng).CurrentRegion
    Region = Rng.value

    For I = 2 To UBound(Region, 1)
        outOutstanding = outOutstanding + Region(I, 6)
        If Region(I, 7) < Date Then
            outOverdue = outOverdue + Region(I, 6)
        End If
   Next I

End Function

Private Function iStatementSupport_FindStatementDataStart() As Range
    Set iStatementSupport_FindStatementDataStart = ExcelTemplate.Sheets(1).Cells.Find("<<Data Starts Here>>")
End Function

Private Function iStatementSupport_GetDataStatement(CustomerID As String, ByRef NRec As Long) As ADODB.Recordset
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        [Tbl_Invoices].[Document_Number]   , "
    SQL = SQL & "        [Tbl_Invoices].[Date]              , "
    SQL = SQL & "        [Tbl_Invoices].[Customer_reference], "
    SQL = SQL & "        [Tbl_Invoices].SONumber            , "
    SQL = SQL & "        [Tbl_Invoices].[Currency]          , "
    SQL = SQL & "        [Tbl_Invoices].Amount              , "
    SQL = SQL & "        [Tbl_Invoices].[Overdue_Date]       "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Invoices] "
    SQL = SQL & "Where "
    SQL = SQL & "        [Tbl_Invoices].[Update_date] = #" & Month(Date) & "/" & Day(Date) & "/" & Year(Date) & "#"
    SQL = SQL & "And     [Tbl_Invoices].[Customer_ID] = '" & CustomerID & "' "
    SQL = SQL & "Order By "
    SQL = SQL & "[Tbl_Invoices].[Overdue_Date]"

    Dim RS As ADODB.Recordset
    Set RS = CurrentProject.Connection.Execute(SQL)
    Set iStatementSupport_GetDataStatement = RS

    SQL = "" ' calculate records retrieved
    SQL = SQL & "Select "
    SQL = SQL & "        Count([Tbl_Invoices].[Overdue_Date]) As [Sum_Overdue_Date] "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Invoices] "
    SQL = SQL & "Where "
    SQL = SQL & "        [Tbl_Invoices].[Update_date] = Date() "
    SQL = SQL & "And     [Tbl_Invoices].[Customer_ID] = '" & CustomerID & "' "
    SQL = SQL & "Order By "
    SQL = SQL & "        Count([Tbl_Invoices].[Overdue_Date]) "
    Dim NumRecRS As Recordset
    Set NumRecRS = CurrentProject.Connection.Execute(SQL)
    NRec = NumRecRS.Fields(0).value
End Function

Private Function iStatementSupport_GetRecordsetStatement(SQLString As String) As ADODB.Recordset
    Set iStatementSupport_GetRecordsetStatement = CurrentProject.Connection.Execute(SQLString)
End Function

Private Function iStatementSupport_GetSQLStatement(CustomerCode As String) As String
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        [Tbl_Invoices].[Document_Number]   , "
    SQL = SQL & "        [Tbl_Invoices].[Date]              , "
    SQL = SQL & "        [Tbl_Invoices].[Customer_reference], "
    SQL = SQL & "        [Tbl_Invoices].SONumber            , "
    SQL = SQL & "        [Tbl_Invoices].[Currency]          , "
    SQL = SQL & "        [Tbl_Invoices].Amount              , "
    SQL = SQL & "        [Tbl_Invoices].[Overdue_Date]      , "
    SQL = SQL & "        [Tbl_Invoices].[Customer_ID]       , "
    SQL = SQL & "        [Tbl_Invoices].[Update_date] "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Invoices] "
    SQL = SQL & "Where "
    SQL = SQL & "        [Tbl_Invoices].[Customer_ID] = '" & CustomerCode & "' "
    SQL = SQL & "And     [Tbl_Invoices].[Update_date] = date "
End Function

Private Function iStatementSupport_GetStatementSavePath() As String
    Dim SQL As String
    Dim RS As DAO.Recordset
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "    TblGeneral.PathExcelDirectory "
    SQL = SQL & "From "
    SQL = SQL & "    TblGeneral "

    Set RS = CurrentDb.OpenRecordset(SQL)
    iStatementSupport_GetStatementSavePath = RS.Fields(0).value
    If Right(iStatementSupport_GetStatementSavePath, 1) <> "\" Then
        iStatementSupport_GetStatementSavePath = iStatementSupport_GetStatementSavePath + "\"
    End If

End Function

Private Function iStatementSupport_GetTemplatePath() As String
    Dim SQL As String
    Dim RS As DAO.Recordset
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "    TblGeneral.PathTemplates "
    SQL = SQL & "From "
    SQL = SQL & "    TblGeneral "

    Set RS = CurrentDb.OpenRecordset(SQL)
    iStatementSupport_GetTemplatePath = RS.Fields(0).value
    If Right(iStatementSupport_GetTemplatePath, 1) <> "\" Then
        iStatementSupport_GetTemplatePath = iStatementSupport_GetTemplatePath + "\"
    End If

    '    iStatementSupport_GetTemplatePath = RS.Fields(0).value
End Function

Private Sub iStatementSupport_Maximize()
    GeneralExcel.WindowState = xlMaximized
End Sub

Private Function iStatementSupport_OpenStatementTemplate(PathTemplate As String) As Boolean
    Set ExcelTemplate = GeneralExcel.Workbooks.Open(PathTemplate & "SimpleStatement.xlsx")
    iStatementSupport_OpenStatementTemplate = True
End Function

Private Function iStatementSupport_SaveStatement(CompleteFileName As String) As Boolean
    ExcelTemplate.SaveAs Filename:=CompleteFileName, FileFormat:=xlWorkbookDefault
    iStatementSupport_SaveStatement = True
End Function

Private Function iStatementSupport_StatementReplace(aWhat As String, aWith As String) As Boolean
    ExcelTemplate.Sheets(1).Cells.Replace what:=aWhat, Replacement:=aWith, MatchCase:=True
    iStatementSupport_StatementReplace = True
End Function

Private Function iStatementSupport_StatementVisible(Visible As Boolean) As Boolean
    GeneralExcel.Visible = Visible
    iStatementSupport_StatementVisible = Visible
End Function
