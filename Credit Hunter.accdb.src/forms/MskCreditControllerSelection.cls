Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command4_Click()
Dim SqlText As String
Dim qdfNew As DAO.QueryDef
Dim I As Integer
With CurrentDb
    On Error Resume Next
    .QueryDefs.Delete ("QueryCashCollectedPerCreditController")
 Rem   SQLString = "SELECT Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Sum([Amount]*[ExchangeRate]) AS [AR Exposure in main currency], Tbl_Invoices.Update_date, Sum(IIf([tbl_invoices.overdue_date]+90<=GetNextMonthEnd(),[amount],0))*[ExchangeRate] AS [Overdue 90+ days], Tbl_Currencies.ExchangeRate, Tbl_Customer_Status.Status, Sum(IIf([tbl_invoices.overdue_date]<=GetNextMonthEnd(),[amount],0))*[Tbl_Currencies].[ExchangeRate] AS [Total Overdue], Tbl_Customers.Country, Sum(IIf([tbl_invoices.overdue_date]<=Now(),[amount],0))*[Tbl_Currencies].[ExchangeRate] AS [Overdue as of Today in main currency], Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Target in Main currency], QueryTotalAlreadyCollectedInEUR.AmountInEUR AS [Already collected in Main currency], [MonthlyTargetInMainCurrency]-[AmountInEUR] AS [Still to be collected], IIf([MonthlyTargetInMainCurrency]=0,0,[AmountInEUR]/[MonthlyTargetInMainCurrency]) AS [% Cash Target Achieved] " & _
        "FROM ((Tbl_Currencies RIGHT JOIN (Tbl_Customers LEFT JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) ON Tbl_Currencies.CurrencyID = Tbl_Invoices.Currency) LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID) LEFT JOIN QueryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyCollectedInEUR.CustomerID " & _
        "GROUP BY Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Tbl_Invoices.Update_date, Tbl_Currencies.ExchangeRate, Tbl_Customer_Status.Status, Tbl_Customers.Country, Tbl_Customers.MonthlyTargetInMainCurrency, QueryTotalAlreadyCollectedInEUR.AmountInEUR, [MonthlyTargetInMainCurrency]-[AmountInEUR], IIf([MonthlyTargetInMainCurrency]=0,0,[AmountInEUR]/[MonthlyTargetInMainCurrency]) " & _
        " HAVING (((Tbl_Customers.Credit_controller)=" & Combo6.Column(1) & ") AND ((Tbl_Invoices.Update_date)=Date()));"

    SQLString = "SELECT Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Sum([Amount]*[ExchangeRate]) AS [AR Exposure in main currency], Tbl_Invoices.Update_date, Sum(IIf([tbl_invoices.overdue_date]+90<=GetNextMonthEnd(),[amount],0))*[ExchangeRate] AS [Overdue 90+ days], Tbl_Currencies.ExchangeRate, Tbl_Customer_Status.Status, Sum(IIf([tbl_invoices.overdue_date]<=GetNextMonthEnd(),[amount],0))*[Tbl_Currencies].[ExchangeRate] AS [Total overdue on fiscal month end ], Tbl_Customers.Country, Sum(IIf([tbl_invoices.overdue_date]<=Now(),[amount],0))*[Tbl_Currencies].[ExchangeRate] AS [Overdue as of Today in main currency], Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Target in Main currency], val(nz([AmountInEUR],0)) AS [Already collected in Main currency], [MonthlyTargetInMainCurrency]-[Already collected in Main currency] AS [Still to be collected in Main currency], " & _
        " IIf([MonthlyTargetInMainCurrency]=0,0,[Already collected in Main currency]/[MonthlyTargetInMainCurrency]) AS [% Cash Target Achieved] " & _
        " FROM ((Tbl_Currencies RIGHT JOIN (Tbl_Customers LEFT JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) ON Tbl_Currencies.CurrencyID = Tbl_Invoices.Currency) LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID) LEFT JOIN QueryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyCollectedInEUR.CustomerID " & _
        " GROUP BY Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Tbl_Invoices.Update_date, Tbl_Currencies.ExchangeRate, Tbl_Customer_Status.Status, Tbl_Customers.Country, Tbl_Customers.MonthlyTargetInMainCurrency, nz([AmountInEUR],0)" & _
        " HAVING (((Tbl_Customers.Credit_controller)=" & Combo6.Column(1) & ") AND ((Tbl_Invoices.Update_date)=Date()));"

    SQLString = "SELECT Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Query11.[AR Exposure in main currency], Query11.Update_date, Query11.[Overdue 90+ days], Query11.ExchangeRate, Tbl_Customer_Status.Status, Query11.[Total overdue on fiscal month end], Tbl_Customers.Country, Query11.[Overdue as of Today in main currency], Tbl_Customers.MonthlyTargetInMainCurrency AS [Monthly Target in Main currency], IIf(IsNull([AmountInEUR]),0,[AmountInEUR]) AS [Already collected in Main currency], [MonthlyTargetInMainCurrency]-[Already collected in Main currency] AS [Still to be collected in Main currency], IIf([MonthlyTargetInMainCurrency]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTargetInMainCurrency],0) AS [% Cash Target Achieved] " & _
        " FROM ((Tbl_Customers LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status = Tbl_Customer_Status.ID) LEFT JOIN QueryTotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyCollectedInEUR.CustomerID) LEFT JOIN Query11 ON Tbl_Customers.Customer_code = Query11.Customer_ID" & _
        " GROUP BY Tbl_Customers.Customer_code, Tbl_Customers.Credit_controller, Tbl_Customers.Name, Query11.[AR Exposure in main currency], Query11.Update_date, Query11.[Overdue 90+ days], Query11.ExchangeRate, Tbl_Customer_Status.Status, Query11.[Total overdue on fiscal month end], Tbl_Customers.Country, Query11.[Overdue as of Today in main currency], Tbl_Customers.MonthlyTargetInMainCurrency, IIf(IsNull([AmountInEUR]),0,[AmountInEUR]), IIf([MonthlyTargetInMainCurrency]<>0,IIf(IsNull([AmountInEUR]),0,[AmountInEUR])*100/[MonthlyTargetInMainCurrency],0)" & _
        "HAVING (((Tbl_Customers.Credit_controller)=" & Combo6.Column(1) & "));"


    Set qdfNew = .CreateQueryDef("QueryCashCollectedPerCreditController", SQLString)
    On Error GoTo 0
    DoEvents

    DoCmd.OutputTo acOutputQuery, "QueryCashCollectedPerCreditController", acFormatXLS, GetPathExcelDirectory() & "QueryCashCollectedPerCreditController.xls", False
    Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "QueryCashCollectedPerCreditController.xls")
    ExcApp.Visible = True


    With ExcDoc.Sheets(1)
        ExcApp.ActiveSheet.Range("A..Z").EntireColumn.AutoFit
        ExcApp.ActiveSheet.Range("A1,D:D,F:F,I:I,K:K,L:L,M:M,N:N").NumberFormat = "#,##0"
        ExcApp.ActiveSheet.Columns("O:O").NumberFormat = "0%"

        I = 1
        While ExcDoc.Sheets(1).Cells(I, 1) <> ""
            I = I + 1
        Wend
        I = I - 1
        ExcDoc.Sheets.Add

        ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "QueryCashCollectedPerCreditCont!R1C1:R" & I & "C15", Version:=xlPivotTableVersion10 _
            ).CreatePivotTable TableDestination:="Sheet1!R3C1", tableName:= _
            "PivotTable1", DefaultVersion:=xlPivotTableVersion10
        ExcDoc.Sheets("Sheet1").Select
        ExcDoc.Sheets(1).Cells(3, 1).Select
        With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("Name")
            .Orientation = xlRowField
            .Position = 1
        End With
        ExcDoc.ActiveSheet.PivotTables("PivotTable1").AddDataField ExcDoc.ActiveSheet.PivotTables( _
            "PivotTable1").PivotFields("Still to be collected in Main currency"), _
            "Count of Still to be collected in Main currency", xlCount
        With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields( _
            "Count of Still to be collected in Main currency")
            .Caption = "Sum of Still to be collected in Main currency"
            .Function = xlSum
        End With
        ExcApp.ActiveSheet.Range("B5").Select
        ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
            xlDescending, "Sum of Still to be collected in Main currency", ExcDoc.ActiveSheet.PivotTables( _
            "PivotTable1").PivotColumnAxis.PivotLines(1), 1

        ExcApp.ActiveSheet.Columns("B:B").Select
        ExcApp.ActiveSheet.Range("B:B").NumberFormat = "#,##0"
        ExcApp.ActiveSheet.Range("B5").Select
        ExcApp.ActiveSheet.PivotTables("PivotTable1").PivotFields("Name").AutoSort _
            xlDescending, "Sum of Still to be collected in Main currency", ExcApp.ActiveSheet.PivotTables( _
            "PivotTable1").PivotColumnAxis.PivotLines(1), 1
    End With
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If

End Sub

Private Sub Form_Load()
Dim CControllerName As String
Dim I As Integer
    Me.Combo6.Visible = DLookup("superuser", "tbl_users", "UserName='" & fOSUserName() & "'") = True
    CControllerName = GetNameCreditController(fOSUserName)
    For I = 0 To Me.Combo6.ListCount - 1
        If Me.Combo6.ItemData(I) = CControllerName Then
            Me.Combo6 = Me.Combo6.ItemData(I)
        End If
    Next

End Sub
