Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command4_Click()
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim RS As Variant
    Dim I As Integer
    Dim S As String
    If CDate(Text0.value) > CDate(Text2.value) Then
        MsgBox "From date has to be before To date. Please check.", vbOKOnly + vbCritical, "Error: "
    Else
  '      CurrentDb.QueryDefs("QueryCashCollected").SQL = "SELECT Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Sum(Tbl_CashCollected.Amount) AS [Amount in EUR], Tbl_Cash_Target.CashTargetInEUR FROM (((Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID) INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID) INNER JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID) INNER JOIN Tbl_Cash_Target ON Tbl_Users.ID = Tbl_Cash_Target.CControllerID WHERE ((Tbl_Customers.Credit_controller) = " & Combo6.Column(1) & ") GROUP BY Tbl_Users.Name, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Tbl_Cash_Target.CashTargetInEUR, Tbl_Customers.RetailOEM HAVING (((Tbl_CashCollected.[Payment Date]) >= #" & format(Text0.Value, "mm/dd/yy") & "# And (Tbl_CashCollected.[Payment Date]) <= #" & format(Text2.Value, "mm/dd/yy") & "#)) ORDER BY Sum(Tbl_CashCollected.Amount) DESC;"

  Rem       S = QueryCashCollectedSQLParser("SELECT Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Sum(Tbl_CashCollected.Amount) AS [Amount in EUR], Tbl_CashCollected.[Currency] , Tbl_CashCollected.[Original amount] FROM (((Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID) INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID) INNER JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID) INNER JOIN Tbl_Cash_Target ON (Tbl_Customers.RetailOEM = Tbl_Cash_Target.Channel) AND (Tbl_Users.ID = Tbl_Cash_Target.CControllerID) GROUP BY Tbl_Customers.RetailOEM, Tbl_Users.Name, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.[Currency] , Tbl_CashCollected.[Original amount] HAVING (((Tbl_CashCollected.[Payment Date]) >= #" & Format(Text0.value, "mm/dd/yy") & "# And (Tbl_CashCollected.[Payment Date]) <= #" & Format(Text2.value, "mm/dd/yy") & "#))" & _
'        " ORDER BY Sum(Tbl_CashCollected.Amount) DESC;" _
 '       , Combo6.Column(0), IIf(Token(Combo6.Column(1), "/", 2) = "All accounts", "", (Token(Combo6.Column(1), "/", 2))))
  '
        S = QueryCashCollectedSQLParser("SELECT Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Sum(Tbl_CashCollected.Amount) AS [Amount in EUR], Tbl_CashCollected.Currency, Tbl_CashCollected.[Original amount], Tbl_Users.ID FROM ((Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID) INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID) INNER JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID GROUP BY Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.Currency, Tbl_CashCollected.[Original amount], Tbl_Users.ID HAVING (((Tbl_CashCollected.[Payment Date])>=#" & Format(Text0.value, "mm/dd/yy") & "# And (Tbl_CashCollected.[Payment Date])<=#" & Format(Text2.value, "mm/dd/yy") & "#)) " & _
        "ORDER BY Sum(Tbl_CashCollected.Amount) DESC;" _
        , Combo6.Column(0), IIf(Token(Combo6.Column(1), "/", 2) = "All accounts", "", (Token(Combo6.Column(1), "/", 2))))


        CurrentDb.QueryDefs("QueryCashCollected").SQL = S

        Set RS = CurrentDb.OpenRecordset("QueryCashCollected")
        RS.MoveLast
        If RS.RecordCount > 0 Then
            DoCmd.OutputTo acOutputQuery, "QueryCashCollected", acFormatXLS, GetPathExcelDirectory() & "QueryCashCollected.xls", False
            Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
            Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "QueryCashCollected.xls")
            ExcApp.Visible = True
            With ExcDoc.Sheets(1)
                .Cells(1, 1) = "Credit Controller"
                .Cells(1, 2) = "Channel"
                .Cells(1, 3) = "Customer name"
                I = 2
                While .Cells(I, 3) <> ""
                    I = I + 1
                Wend
                If I > 2 Then
                    .Cells(I + 2, 4) = "Total collected"
                    .Cells(I + 2, 5) = "=Sum(E" & I - 1 & ":E2)"

                    I = I + 2
              '      .Cells(i + 2, 5) = rs.Fields("CashTargetInEUR")
               '     .Columns("E:E").NumberFormat = "#,##0.00"
                '    .Cells(i + 2, 4) = "Cash Target"
                    With ExcDoc
                        .Sheets.Add
                        ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                            "QueryCashCollected!R1C1:R" & I - 1 & "C5", Version:=xlPivotTableVersion10). _
                            CreatePivotTable TableDestination:="Sheet1!R3C1", tableName:="PivotTable1" _
                            , DefaultVersion:=xlPivotTableVersion10
                        .Sheets("Sheet1").Select
'                        .Cells(3, 1).Select
                        .ActiveSheet.PivotTables("PivotTable1").AddDataField .ActiveSheet.PivotTables( _
                            "PivotTable1").PivotFields("Amount in EUR"), "Sum of Amount in EUR", xlSum
                        With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("Customer name")
                            .Orientation = xlRowField
                            .Position = 1
                        End With
                        .ActiveSheet.Range("B7").Select
                        .ActiveSheet.PivotTables("PivotTable1").PivotFields("Customer name"). _
                            AutoSort xlDescending, "Sum of Amount in EUR", .ActiveSheet.PivotTables( _
                            "PivotTable1").PivotColumnAxis.PivotLines(1), 1
                        With .ActiveSheet.PivotTables("PivotTable1").PivotFields("Payment Date")
                            .Orientation = xlRowField
                            .Position = 2
                        End With
                        .ActiveSheet.Range("C8").Select
                        .ActiveSheet.PivotTables("PivotTable1").PivotFields("Payment Date").AutoSort _
                            xlDescending, "Sum of Amount in EUR", .ActiveSheet.PivotTables("PivotTable1"). _
                            PivotColumnAxis.PivotLines(1), 1
                        .ActiveSheet.Columns("C:C").NumberFormat = "#,##0.00"

                        .Sheets("QueryCashCollected").Select
                        ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                            "QueryCashCollected!R1C1:R" & I - 1 & "C5", Version:=xlPivotTableVersion10). _
                            CreatePivotTable TableDestination:="Sheet1!R3C5", tableName:="PivotTable4" _
                            , DefaultVersion:=xlPivotTableVersion10
                        .Sheets("Sheet1").Select
                        .ActiveSheet.PivotTables("PivotTable4").AddDataField .ActiveSheet.PivotTables( _
                            "PivotTable4").PivotFields("Amount in EUR"), "Sum of Amount in EUR", xlSum
                        With .ActiveSheet.PivotTables("PivotTable4").PivotFields("Payment Date")
                            .Orientation = xlRowField
                            .Position = 1
                        End With
                        .ActiveSheet.Range("E8").Select
                        .ActiveSheet.PivotTables("PivotTable4").PivotFields("Payment Date").AutoSort _
                            xlAscending, "Payment Date"
                        .ActiveSheet.Columns("F:F").NumberFormat = "#,##0.00"
                        .ActiveSheet.Range("A1").Select


                        ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                            "QueryCashCollected!R1C1:R" & I - 1 & "C7", Version:=xlPivotTableVersion10). _
                            CreatePivotTable TableDestination:="Sheet1!R6C8", tableName:="PivotTable5" _
                            , DefaultVersion:=xlPivotTableVersion10

                        With .ActiveSheet.PivotTables("PivotTable5").PivotFields("Credit Controller")
                            .Orientation = xlPageField
                            .Position = 1
                        End With
                        With .ActiveSheet.PivotTables("PivotTable5").PivotFields("Channel")
                            .Orientation = xlPageField
                            .Position = 1
                        End With

                        .ActiveSheet.PivotTables("PivotTable5").AddDataField .ActiveSheet.PivotTables( _
                            "PivotTable5").PivotFields("Original amount"), "Sum of Original amount", xlSum

                        With .ActiveSheet.PivotTables("PivotTable5").PivotFields("Currency")
                            .Orientation = xlRowField
                            .Position = 1
                        End With
                        .ActiveSheet.Columns("I:I").NumberFormat = "#,##0.00"
                        .ActiveSheet.Columns("I:I").AutoFit
                        .ActiveSheet.PivotTables("PivotTable5").ColumnGrand = False

                        .Sheets(2).Select
                        .Sheets(2).Name = "Query Collected - Details"

                        .Sheets(1).Select
                        .Sheets(1).Name = "Query Collected - Pivots"

                        ExcApp.ActiveWindow.TabRatio = 0.409
                        ExcApp.ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
                        .Sheets(1).Select
                    End With
                Else
                    MsgBox "No payments received between the dates " & Format(Text0.value, "dd-mmm-yy") & " and " & Format(Text0.value, "dd-mmm-yy") & " .", vbOKOnly, "Error: "
                End If
            End With
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rst2 As Variant
    Dim PrevCreditControllerID As Integer
'    Combo6.visible = DLookup("[Superuser]", "Tbl_Users", "ID=" & GetNumCreditController(fOSUserName()))
 '   Combo6.Value = Combo6.ItemData(0)
    Combo6.RowSource = ""
    If DLookup("Superuser", "Tbl_Users", "ID=" & GetNumCreditController(fOSUserName())) Then
        Combo6.RowSource = "0;All accounts;"
        Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Customers.RetailOEM FROM Tbl_Customers GROUP BY Tbl_Customers.RetailOEM;")
        rst2.MoveFirst
        While Not rst2.EOF
            Combo6.RowSource = Combo6.RowSource & "0;All accounts/" & rst2.Fields("retailOEM") & ";"
            rst2.MoveNext
        Wend
        Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Users.ID, Tbl_Users.Name, Tbl_Customers.RetailOEM FROM Tbl_Users INNER JOIN Tbl_Customers ON Tbl_Users.ID = Tbl_Customers.Credit_controller GROUP BY Tbl_Users.ID, Tbl_Users.Name, Tbl_Customers.RetailOEM;")

    Else
        Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Users.ID, Tbl_Users.Name, Tbl_Customers.RetailOEM FROM Tbl_Users INNER JOIN Tbl_Customers ON Tbl_Users.ID = Tbl_Customers.Credit_controller GROUP BY Tbl_Users.ID, Tbl_Users.Name, Tbl_Customers.RetailOEM HAVING (((Tbl_Users.ID)=" & GetNumCreditController(fOSUserName()) & "));")
    End If


    Combo6.Visible = rst2.RecordCount > 1
    PrevCreditControllerID = 0
    While Not rst2.EOF
        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Customers.RetailOEM, Tbl_Customers.Credit_controller FROM Tbl_Customers GROUP BY Tbl_Customers.RetailOEM, Tbl_Customers.Credit_controller HAVING (((Tbl_Customers.Credit_controller)=" & rst2.Fields("ID") & "));")
        rst.MoveLast
        If (rst.RecordCount > 1) And (PrevCreditControllerID <> rst2.Fields("ID")) Then
            PrevCreditControllerID = rst2.Fields("ID")
            Combo6.RowSource = Combo6.RowSource & rst2.Fields("ID") & ";" & rst2.Fields("Name") & "/All accounts" & ";"
        End If
        Combo6.RowSource = Combo6.RowSource & rst2.Fields("ID") & ";" & rst2.Fields("Name") & "/" & rst2.Fields("RetailOEM") & ";"
        rst2.MoveNext
    Wend
    Combo6.value = Me.Combo6.ItemData(0)

    Text0.value = DateAdd("d", 1, DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd <#" & Format(Date, "mm/dd/yy") & "#"))
    Text2.value = DMin("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd >#" & Format(DateAdd("d", 1, CDate(Text0.value)), "mm/dd/yy") & "#")

    If Me.OpenArgs <> "" Then
        Me.Combo6.value = Me.OpenArgs
    End If
End Sub
