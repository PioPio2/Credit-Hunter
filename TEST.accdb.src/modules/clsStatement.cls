VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements IStatement
Private CompletePath As String

Private Sub CreateExcel()
    If GeneralExcel Is Nothing Then
        Set GeneralExcel = New Excel.Application
    End If
End Sub

Private Sub TurnOffCalculations()
    If Not GeneralExcel Is Nothing Then
        GeneralExcel.Calculation = xlCalculationManual
        GeneralExcel.ScreenUpdating = False
        GeneralExcel.EnableEvents = False
    End If
End Sub

Private Sub TurnOnCalculations()
    If Not GeneralExcel Is Nothing Then
        GeneralExcel.Calculation = xlCalculationAutomatic
        GeneralExcel.ScreenUpdating = True
        GeneralExcel.EnableEvents = True
    End If
End Sub

'this is the former version. Get rid as soon as the main function is operating
Private Function FormerExcelStatement(Customer As Recordset, CurrencyTab As Variant, rstbanks As Recordset, Optional Monthend As Date, Optional CloseStatement As Boolean) As String
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim ExeWksht As Excel.Worksheet
    Dim Typecurrency, FileName As String
    Dim NextEmptyColumn, sheet, Currinv, heet, c, r, row, col, StartingDataRow, StartingDataRow2, StartingDataRow3, InvoiceLine, COLWIDTH As Integer
    Dim ColCustInvn, PullTicketN, OriginalAmount As Integer
    Dim MonthEndAmount, Current, O31Days, O3160Days, O61days As Currency
    Dim QuerySpec As Recordset
    Dim DirSave As String
    Dim aa, bb, ExtraColumns As Integer
    Dim recc As DAO.Recordset
    Dim DocumentTypes As Variant
    Dim DocumentsToBeErased As Variant
    Dim AddColumnCustomsInvoicenum, AddPullTicketn, AddOriginalTransactionAmount As Boolean

    Set ExcApp = CreateObject("Excel.Application")    'apre il modello di Excel
    'Set ExcDoc = ExcApp.Workbooks.Add
    Dim PathTemplate As String
    PathTemplate = DLookup("PathTemplates", "tblGeneral")
    PathTemplate = PathTemplate & "\"
    Set ExcDoc = ExcApp.Workbooks.Open(PathTemplate & "SimpleStatement.xlsx")
    ExcApp.Visible = True
    InvoiceLine = 1
    sheet = 1
    aa = 0

    'conta quanti saranno le valute nell'estratto conto alla fine(tabs)
    If CurrencyTab.Controls.item("Pagina88").Visible Then
        aa = aa + 1
    End If
    If CurrencyTab.Controls.item("Pagina89").Visible Then
        aa = aa + 1
    End If
    If CurrencyTab.Controls.item("Pagina90").Visible Then
        aa = aa + 1
    End If
    If CurrencyTab.Controls.item("Pagina91").Visible Then
        aa = aa + 1
    End If
    If CurrencyTab.Controls.item("Pagina92").Visible Then
        aa = aa + 1
    End If

    Set QuerySpec = New Recordset
    With QuerySpec
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_queries", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    With rstbanks
        .MoveFirst
        .Find ("Country='" & Customer.Fields("Country") & "'")
    End With

    For bb = 1 To aa                             'ripete per ogni tab (valuta)
        Select Case bb                           'mette in recc la porzione di e/c corrispondente ad alla valuta che si sta esaminando
        Case 1
            Set recc = CurrencyTab.Controls.item("Pagina88").Controls.item(0).Form.RecordsetClone
        Case 2
            Set recc = CurrencyTab.Controls.item("Pagina89").Controls.item(0).Form.RecordsetClone
        Case 3
            Set recc = CurrencyTab.Controls.item("Pagina90").Controls.item(0).Form.RecordsetClone
        Case 4
            Set recc = CurrencyTab.Controls.item("Pagina91").Controls.item(0).Form.RecordsetClone
        Case 5
            Set recc = CurrencyTab.Controls.item("Pagina92").Controls.item(0).Form.RecordsetClone
        End Select
        recc.MoveFirst
        Currinv = 0

        Set DocumentTypes = New Recordset
        With DocumentTypes
            Set DocumentTypes = CurrentDb.OpenRecordset("SELECT Tbl_Types.ID, Tbl_Types.Descripition FROM Tbl_Types;")
        End With

        MonthEndAmount = 0
        While (Not recc.EOF)
            AddColumnCustomsInvoicenum = Customer("PullTicketNumberToBePrinted")
            If AddColumnCustomsInvoicenum Then
                ColCustInvn = 2
            End If


            AddPullTicketn = Customer("FacturaNumberToBePrinted")
            If AddPullTicketn Then
                If AddColumnCustomsInvoicenum Then
                    PullTicketN = 3
                Else
                    PullTicketN = 2
                End If
            End If

            AddOriginalTransactionAmount = Customer("OriginalInvoiceAmountToBePrinted")
            If AddOriginalTransactionAmount Then
                OriginalAmount = 9
                If AddPullTicketn = False Then
                    OriginalAmount = OriginalAmount - 1
                End If
                If AddColumnCustomsInvoicenum = False Then
                    OriginalAmount = OriginalAmount - 1
                End If
            End If
            If ExcDoc.Worksheets.Count < sheet Then
                ExcDoc.Worksheets.Add
            End If
            ExcApp.Worksheets(sheet).Select
            row = 9
            col = 1
            With ExcApp
                Rem ExcApp.ActiveSheet.Insert.Picture (GetPathlogo())
                Rem        .ActiveSheet.Pictures.Insert (GetPathlogo())
                '        .ActiveSheet.Rows("1:1").RowHeight = 140

                '       .Cells(3, 1) = "STATEMENT OF ACCOUNT"
                '      .Cells(3, 1).Select
                '     .Cells(3, 1).Font.Bold = True
                '    .Cells(3, 1).Font.Italic = True
                '   .Cells(3, 1).Font.Size = 16
                '        ExcDoc.Worksheets(sheet).Range("A3:G3").HorizontalAlignment = xlCenterAcrossSelection
                '       .Cells(row, col) = Customer.Fields("OWN_company")
                '      .Cells(row, col).Font.Bold = True
                'row = row + 1
                'r = row
                'c = col
                Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details1")
                '.Cells(row, col).Font.Bold = True
                '        row = row + 1
                Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details2")
                '       .Cells(row, col).Font.Bold = True
                '      row = row + 1
                Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details3")
                '     .Cells(row, col).Font.Bold = True
                '    row = row + 1
                Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details4")
                '   .Cells(row, col).Font.Bold = True
                '  row = row + 2
                ' col = 1
                .Cells(row, col) = Customer.Fields("Name")
                .Cells(row, col).Font.Bold = True
                .Cells(row, col + 4) = "Customer ID: " & Customer.Fields("customer_code")
                .Cells(row, col + 4).Font.Bold = True

                row = row + 1
                .Cells(row, col) = Customer.Fields("Address")
                '        .Cells(row, col).Font.Bold = True
                row = row + 1
                .Cells(row, col) = Customer.Fields("Address2")
                '       .Cells(row, col).Font.Bold = True
                If Customer.Fields("Address3") <> "" Then
                    row = row + 1
                    .Cells(row, col) = Customer.Fields("Address3")
                    '          .Cells(row, col).Font.Bold = True
                End If
                If Customer.Fields("Address4") <> "" Then
                    row = row + 1
                    .Cells(row, col) = Customer.Fields("Address4")
                    '         .Cells(row, col).Font.Bold = True
                End If
                row = row + 1

                .Cells(row, col) = Customer.Fields("Country")

                .Cells(row, col + 4) = "Date: " & Format(Date, "dd-mmm-yyyy")
                .Cells(row, col + 4).Font.Bold = True

                row = row + 4
                StartingDataRow = row
                StartingDataRow2 = row
                StartingDataRow3 = row
                '        .Cells(row, col + 0) = "Invoice number"
                '       .Cells(row, col + 1) = "Date"
                '      .Cells(row, col + 1).HorizontalAlignment = xlCenter

                '     .Cells(row, col + 2) = "Type"
                '    .Cells(row, col + 3) = "Reference"
                '   .Cells(row, col + 4) = "SO number"
                '  .Cells(row, col + 5) = "Amount due"
                '        .Cells(row, col + 6).HorizontalAlignment = xlRight
                '       .Cells(row, col + 6) = "Due date"
                '      .Cells(row, col + 6).HorizontalAlignment = xlRight
                If Customer.Fields("StatementForm") > 0 Then
                    If Customer.Fields("StatementForm") = 1 Then
                        .Cells(row, col + 7) = "Queries"
                        .Cells(row, col + 7).HorizontalAlignment = xlLeft
                    ElseIf Customer.Fields("StatementForm") = 3 Then
                        .Cells(row, col + 7) = "Queries"
                        .Cells(row, col + 7).HorizontalAlignment = xlLeft
                        .Cells(row, col + 8) = "Notes"
                        .Cells(row, col + 8).HorizontalAlignment = xlLeft

                    ElseIf Customer.Fields("StatementForm") = 4 Then
                        .Cells(row, col + 7) = "Notes"
                        .Cells(row, col + 7).HorizontalAlignment = xlLeft
                    End If
                End If

                NextEmptyColumn = 1
                While .Cells(row, NextEmptyColumn) <> ""
                    NextEmptyColumn = NextEmptyColumn + 1
                Wend
                If Monthend <> "00:00:00" Then
                    .Cells(row, NextEmptyColumn) = "Total amount to be paid by month's end (" & Format(Monthend, "dd mmmm  yyyy") & ")"
                    .Cells(row, NextEmptyColumn).HorizontalAlignment = xlCenter
                    .Cells(row, NextEmptyColumn).WrapText = True
                End If

                'row = row + 1
                Typecurrency = recc("Tbl_Invoices.Currency")
                If Not (rstbanks.EOF) Then
                    If Typecurrency = "EUR" Then
                        '                .Cells(r, c) = rstbanks.Fields("EURLine1")
                        '               .Cells(r + 1, c) = rstbanks.Fields("EURLine2")
                        '              .Cells(r + 2, c) = rstbanks.Fields("EURLine3")
                        '             .Cells(r + 3, c) = rstbanks.Fields("EURLine4")
                    ElseIf Typecurrency = "USD" Then
                        '            .Cells(r, c) = rstbanks.Fields("USDLine1")
                        '           .Cells(r + 1, c) = rstbanks.Fields("USDLine2")
                        '          .Cells(r + 2, c) = rstbanks.Fields("USDLine3")
                        '         .Cells(r + 3, c) = rstbanks.Fields("USDLine4")
                    ElseIf Typecurrency = "GBP" Then
                        '        .Cells(r, c) = rstbanks.Fields("GBPLine1")
                        '       .Cells(r + 1, c) = rstbanks.Fields("GBPLine2")
                        '      .Cells(r + 2, c) = rstbanks.Fields("GBPLine3")
                        '     .Cells(r + 3, c) = rstbanks.Fields("GBPLine4")
                    Else
                        '    .Cells(r, c) = ""
                        '   .Cells(r + 1, c) = ""
                        '  .Cells(r + 2, c) = ""
                        ' .Cells(r + 3, c) = ""
                    End If
                End If
                Current = 0
                O31Days = 0
                O3160Days = 0
                O61days = 0
                InvoiceLine = row

                Set DocumentsToBeErased = New Recordset
                With DocumentsToBeErased
                    Set DocumentsToBeErased = CurrentDb.OpenRecordset("SELECT Tbl_DocumentsToBeErased.CustomerID, Tbl_DocumentsToBeErased.DocumentType FROM Tbl_DocumentsToBeErased WHERE (((Tbl_DocumentsToBeErased.CustomerID)='" & Customer.Fields("customer_code") & "'));")
                End With

                On Error GoTo avanti
                While (Not recc.EOF)
                    DocumentsToBeErased.FindFirst "Tbl_DocumentsToBeErased.DocumentType=" & ((recc.Fields("Type")))
                    If (DocumentsToBeErased.RecordCount = 0) Or (DocumentsToBeErased.RecordCount <> 0 And DocumentsToBeErased.NoMatch = True) Then
                        .Cells(row, col + 0) = recc.Fields("document_number")
                        .Cells(row, col + 1) = Format(recc.Fields("date"), "dd/mmm/yyyy")
                        DocumentTypes.MoveFirst
                        '            DocumentTypes.Move ((recc.Fields("Type"))) - 1
                        '.Cells(row, col + 2) = DocumentTypes.Fields("Descripition")
                        .Cells(row, col + 2) = recc.Fields("Customer_reference")
                        .Cells(row, col + 3) = recc.Fields("SONumber")
                        .Cells(row, col + 4) = recc.Fields("Tbl_Invoices.Currency")
                        .Cells(row, col + 5) = Format(recc.Fields("amount"), "##,##0.00")
                        .Cells(row, col + 6) = Format(recc.Fields("Tbl_Invoices.Overdue_Date"), "dd/mmm/yyyy")
                        If Not IsNull(recc.Fields("Query")) Then
                            QuerySpec.MoveFirst
                            QuerySpec.Find ("ID=" & recc.Fields("Query"))
                        End If
                        If Customer.Fields("StatementForm") > 0 Then
                            QuerySpec.MoveFirst
                            If Not IsNull(recc.Fields("Query")) Then
                                QuerySpec.Find ("ID=" & recc.Fields("Query"))
                                If Not QuerySpec.EOF Then
                                    If Customer.Fields("StatementForm") = 1 Then
                                        If recc.Fields("QueryToBePrinted") = True Then
                                            .Cells(row, col + 7) = QuerySpec.Fields("Query")
                                        End If
                                    ElseIf Customer.Fields("StatementForm") = 3 Then
                                        If recc.Fields("QueryToBePrinted") = True Then
                                            .Cells(row, col + 7) = QuerySpec.Fields("Query")
                                        End If
                                        If Not IsNull(recc.Fields("Memo")) Then
                                            If recc.Fields("QueryToBePrinted") = True Then
                                                .Cells(row, col + 8) = recc.Fields("Memo")
                                            End If
                                        End If
                                    ElseIf Customer.Fields("StatementForm") = 4 Then
                                        If Not IsNull(recc.Fields("Memo")) Then
                                            If recc.Fields("QueryToBePrinted") = True Then
                                                .Cells(row, col + 7) = recc.Fields("Memo")
                                            End If
                                        End If
                                    End If
                                    If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And _
                                                                                              ((QuerySpec.Fields("InvoiceToBePaid"))) Then
                                        .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                        .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                    End If
                                Else
                                    If recc.Fields("Tbl_Invoices.overdue_date") <= Monthend Then
                                        .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                        .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                    End If
                                End If
                            Else
                                If (Not (QuerySpec.EOF)) Then
                                    'If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And (QuerySpec.Fields("InvoiceToBePaid")) Then
                                    '                        If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) Then
                                    '                           .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                    '                          .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                    '                     End If
                                    '                Else
                                    If (recc.Fields("Tbl_Invoices.Overdue_Date") <= Monthend) Then
                                        .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                        .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                    End If
                                End If

                            End If
                        Else
                            If (Not (QuerySpec.EOF)) Then
                                'If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And (QuerySpec.Fields("InvoiceToBePaid")) Then
                                '                        If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) Then
                                '                           .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                '                          .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                '                     End If
                                '                Else
                                If (recc.Fields("Tbl_Invoices.Overdue_Date") <= Monthend) Then
                                    .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                    .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                                End If
                            End If
                        End If

                        If recc.Fields("Tbl_Invoices.Overdue_Date") > Date Then
                            Current = Current + recc.Fields("amount")
                        ElseIf DateAdd("d", 61, recc.Fields("Tbl_Invoices.Overdue_Date")) <= Date Then
                            O61days = O61days + recc.Fields("amount")
                        ElseIf DateAdd("d", 31, recc.Fields("Tbl_Invoices.overdue_date")) <= Date Then
                            O3160Days = O3160Days + recc.Fields("amount")
                        Else
                            O31Days = O31Days + recc.Fields("amount")

                        End If
                        recc.MoveNext
                        row = row + 1
                    Else
                        recc.MoveNext
                    End If
                Wend
avanti:
                row = row + 2
                '            .Cells(row, 1) = "Balance due: " & Typecurrency
                '            .Cells(row, 6).NumberFormat = "##,##0.00"
                ExcDoc.Worksheets(sheet).Cells(row, 4) = "Total Outstanding"
                ExcDoc.Worksheets(sheet).Cells(row, 6) = "=Sum(F1:F" & row - 1 & ")"
                'row = row + 2
                'ExcDoc.Worksheets(sheet).Cells(row, 4) = "Total overdue "


                row = row + 3
                '            .Cells(row, 3) = "1-30 Days"
                '           .Cells(row, 3).Font.Bold = True
                '          .Cells(row, 3).HorizontalAlignment = xlCenter
                '         .Cells(row, 4) = "31-60 Days"
                '        .Cells(row, 4).Font.Bold = True
                '       .Cells(row, 4).HorizontalAlignment = xlCenter
                '      .Cells(row, 5) = "61+ Days"
                '     .Cells(row, 5).Font.Bold = True
                '    .Cells(row, 5).HorizontalAlignment = xlCenter
                row = row + 1
                '            .Cells(row, 2) = "Current"
                '           .Cells(row, 2).Font.Bold = True
                '          .Cells(row, 2).HorizontalAlignment = xlCenter
                '         .Cells(row, 3) = "Past due"
                '        .Cells(row, 3).Font.Bold = True
                '       .Cells(row, 3).HorizontalAlignment = xlCenter
                '      .Cells(row, 4) = "Past due"
                '     .Cells(row, 4).Font.Bold = True
                '            .Cells(row, 4).HorizontalAlignment = xlCenter
                '           .Cells(row, 5) = "Past due"
                '          .Cells(row, 5).Font.Bold = True
                '         .Cells(row, 5).HorizontalAlignment = xlCenter
                row = row + 2


                '            .Cells(row, 1) = "Customer Bill-To Total:"
                '           .Cells(row, 1).Font.Bold = True
                '          .Cells(row, 2) = Format(Current, "##,##0.00")
                '         .Cells(row, 2).Font.Bold = True
                '        .Cells(row, 3) = Format(O31Days, "##,##0.00")
                '       .Cells(row, 3).Font.Bold = True
                '      .Cells(row, 4) = Format(O3160Days, "##,##0.00")
                '            .Cells(row, 4).Font.Bold = True
                '           .Cells(row, 5) = Format(O61days, "##,##0.00")
                '          .Cells(row, 5).Font.Bold = True
                '            ExcDoc.Worksheets(sheet).Range("A" & row - 6 & ":E" & row).Columns.AutoFit
                '           ExcDoc.Worksheets(sheet).Columns("b:c").AutoFit
                '          ExcDoc.Worksheets(sheet).Columns("e:f").AutoFit
                '            If (Customer.Fields("StatementForm") = 3) Or (Customer.Fields("StatementForm") = 4) Then
                '               .Range("I:I").ColumnWidth = 50
                '              .Range("I:I").WrapText = True
                '         End If

                '            If Monthend <> "00:00:00" Then
                '               .Cells(row + 2, 1) = "Total amount to be paid by month's end (" & Format(Monthend, "dd mmmm  yyyy") & ")"
                '              .Cells(row + 2, 1).Font.Bold = True
                '             .Cells(row + 2, NextEmptyColumn) = "=Sum(" & Chr(65 + NextEmptyColumn - 1) & "1:" & Chr(65 + NextEmptyColumn - 1) & row - 1 & ")"
                '            .Cells(row + 2, NextEmptyColumn).NumberFormat = "##,##0.00"
                '           .Range("A" & Trim(Str(row + 2)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row + 2))).Interior.color = vbRed
                '          .Cells(row + 2, NextEmptyColumn).Font.Bold = True
                '         .Cells(row + 2, NextEmptyColumn).Columns.AutoFit
                '    End If

                '            ExcDoc.Worksheets(sheet).Columns("g:h").AutoFit
                '           If ExcDoc.Worksheets(sheet).Columns("g").ColumnWidth < 16 Then
                '              ExcDoc.Worksheets(sheet).Columns("g").ColumnWidth = 16
                '         End If

                Rem            .ActiveSheet.Shapes("Picture 1").Select

                '           With ExcDoc.Worksheets(sheet).PageSetup
                '                .Zoom = False
                '               .FitToPagesWide = 1
                '              .FitToPagesTall = 1000
                '         End With

                Rem            ExcApp.Selection.ShapeRange.IncrementLeft (-115)

                '            If Customer.Fields("StatementForm") = 1 Then
                '               ExtraColumns = 1
                '          ElseIf Customer.Fields("StatementForm") = 3 Then
                '             ExtraColumns = 2
                '        End If
                Rem            If .ActiveSheet.Shapes("Picture 1").Width < ExcDoc.Worksheets(sheet).Columns("A:" & Chr(Asc("G") + ExtraColumns)).Width Then
                Rem             ExcApp.Selection.ShapeRange.IncrementLeft (ExcDoc.Worksheets(sheet).Columns("A:" & Chr(Asc("G") + ExtraColumns)).Width - .ActiveSheet.Shapes("Picture 1").Width) / 2
                Rem      End If
                Rem            ExcApp.Selection.ShapeRange.IncrementTop -5000
                Rem         ExcApp.Selection.ShapeRange.IncrementTop 5
                row = row
                row = row
            End With

            '    ExcDoc.Worksheets(sheet).Name = Typecurrency

            '    If AddColumnCustomsInvoicenum Then
            '       ExcDoc.Worksheets(sheet).Columns("B").Insert Shift:=xlToRight
            '       recc.MoveFirst
            '      ExcDoc.Worksheets(sheet).Cells(StartingDataRow, 2) = "Factura n#"
            '     StartingDataRow = StartingDataRow + 1
            '        While Not recc.EOF
            '           ExcDoc.Worksheets(Sheet).Cells(StartingDataRow, 2) = recc.Fields("CustomsInvoiceNumber")
            '          ExcDoc.Worksheets(Sheet).Cells(StartingDataRow, 2).HorizontalAlignment = xlCenter
            '         StartingDataRow = StartingDataRow + 1
            '        recc.MoveNext
            '   Wend
            '        ExcDoc.Worksheets(Sheet).Columns("B:B").AutoFit
            '    End If

            '   If AddPullTicketn Then
            '      ExcDoc.Worksheets(sheet).Columns("B").Insert Shift:=xlToRight
            '        recc.MoveFirst
            '     ExcDoc.Worksheets(sheet).Cells(StartingDataRow2, 2) = "Pull Ticket n#"
            'End If

            '    If AddOriginalTransactionAmount Or AddPullTicketn Or AddColumnCustomsInvoicenum Then
            '       ExcDoc.Worksheets(sheet).Columns("" & Chr(64 + OriginalAmount) & "").Insert Shift:=xlToRight
            '      recc.MoveFirst
            '     ExcDoc.Worksheets(sheet).Cells(StartingDataRow3, OriginalAmount) = "Original Transaction Amount"
            '       StartingDataRow3 = StartingDataRow3 + 1
            '        StartingDataRow = 18


            '        While Not recc.EOF
            '           If AddColumnCustomsInvoicenum Then
            '              ExcDoc.Worksheets(sheet).Cells(StartingDataRow, ColCustInvn) = recc.Fields("CustomsInvoiceNumber")
            '             ExcDoc.Worksheets(sheet).Cells(StartingDataRow, ColCustInvn).HorizontalAlignment = xlCenter
            '        End If
            '
            '      If AddPullTicketn Then
            '         ExcDoc.Worksheets(sheet).Cells(StartingDataRow, PullTicketN) = recc.Fields("[PullTicketN#]")
            '        ExcDoc.Worksheets(sheet).Cells(StartingDataRow, PullTicketN).HorizontalAlignment = xlCenter
            '   End If

            '            If AddOriginalTransactionAmount Then
            '               ExcDoc.Worksheets(sheet).Cells(StartingDataRow, OriginalAmount) = recc.Fields("[OriginalAmount]")
            '              ExcDoc.Worksheets(sheet).Cells(StartingDataRow, OriginalAmount).HorizontalAlignment = xlRight
            '         End If
            '        StartingDataRow = StartingDataRow + 1
            '       recc.MoveNext
            '  Wend
            ' ExcDoc.Worksheets(sheet).Columns("" & Chr(65 + col) & ":" & Chr(65 + col) & "").AutoFit
            '    End If
            '   sheet = sheet + 1
        Wend
    Next bb
    row = row

    'Call AddPaymentsReceived(ExcApp, ExcDoc, Customer.Fields("Customer_code"))
    'ExcApp.ActiveSheet.PageSetup.CenterFooter = "Pag. &P of &N"

    DirSave = DLookup("PathStatemets", "TblGeneral")
    DirSave = Replace(DirSave, "*username*", fOSUserName())
    'If Dir(Dirsave, 16) = "" Then
    '   MkDir (Dirsave)
    'End If

    'Dirsave = "C:\Users\" & fOSUserName() & "\Statements\" & Customer.Fields("Country") & "\"
    'If Dir(Dirsave, 16) = "" Then
    '   MkDir (Dirsave)
    'End If

    'Dirsave = "C:\Users\" & fOSUserName() & "\Statements\" & Customer.Fields("Country") & "\" & Format((Date), "dd-mm-yyyy") & "\"
    'If Dir(Dirsave, 16) = "" Then
    '   MkDir (Dirsave)
    'End If

    FileName = Customer.Fields("Name")
    FileName = NormalizeFileName(FileName)

    FormerExcelStatement = DirSave & FileName & " - " & Format((Now), "dd mmm yyyy - hh.mm.ss") & ".xlsx"
    ExcApp.ActiveWorkbook.SaveAs FileName:=FormerExcelStatement, FileFormat:=xlWorkbookDefault
    If CloseStatement = True Then
        ExcApp.Quit
        Set ExcApp = Nothing
        Set ExcDoc = Nothing
        DocumentTypes.Close
        DocumentTypes = Null
        DocumentsToBeErased.Close
    End If
    QuerySpec.Close
End Function



Public Function IStatement_CreateStatement(Dependency As iStatementSupport, Header As clsCustomerHeader, ByRef outOutstanding As Currency, ByRef outOverdue As Currency, outCurrency As String, Optional CloseStatement As Boolean) As String
    CreateExcel                                  ' create Excel.exe instance if necessary

    Dim PathTemplate As String
    PathTemplate = Dependency.GetTemplatePath

    Dependency.OpenStatementTemplate (PathTemplate)
    Call TurnOffCalculations
    Dependency.StatementVisible (True)

    ' TODO develop here the part of statement with multiple currencies --> one tab for each currency

    Call Dependency.StatementReplace("<<Customer Name>>", Header.getCustomerName)    'replace tags with real customer name
    Call Dependency.StatementReplace("<<Address>>", Header.getAddress)    'replace tags with real customer address
    Call Dependency.StatementReplace("<<TodaysDate>>", Format(Date, "dd-mmm-yyyy"))    'replace tags with real today's date

    Dim Rng As Range
    Set Rng = Dependency.FindStatementDataStart  ' find line where the tag <<Data Starts Here>> is. This is where the data should start in statement template

    Dim RS As ADODB.Recordset
    Dim NRec As Long
    Set RS = Dependency.GetDataStatement(Header.getCustomerID, NRec)

    outOutstanding = 0
    outOverdue = 0
    Call Dependency.CopyStatementData(Rng, RS, outOutstanding, outOverdue, NRec)    ' copy data in the statement template
    RS.MoveLast
    RS.MoveFirst
    outCurrency = RS.Fields("Currency").value

    Call TurnOnCalculations

    Dim FileName As String
    FileName = Header.getCustomerName
    FileName = NormalizeFileName(FileName)

    Dim DirSave As String
    DirSave = Dependency.GetStatementSavePath
    CompletePath = DirSave & FileName & " - " & Format((Now), "dd mmm yyyy - hh.mm.ss") & ".xlsx"
    Call Dependency.SaveStatement(CompletePath)
    If CloseStatement = True Then
        Call Dependency.CloseStatement
    Else
        Dependency.Maximize
    End If

    IStatement_CreateStatement = True
End Function

Private Function IStatement_GetPathStatement() As String
    IStatement_GetPathStatement = CompletePath
End Function
