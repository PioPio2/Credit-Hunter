Attribute VB_Name = "Utility3"
Option Compare Database
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4                    'stop playing
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
                                (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Dim SndFile As String
Dim wFlags As Double
Dim PlayIt

Sub MakeGeneralQueryLogFile()
    Dim RS As Variant
    Dim MyXl As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row As Integer
    Dim FormulaRowStart As Integer

    Dim CountryName As String

    Set RS = CurrentDb.OpenRecordset("SELECT Tbl_AdditionalQueryData.Query_date, Tbl_Areas.Area, Tbl_Customers.Country, Tbl_Customers.Name, Tbl_Customers.Customer_code, Tbl_Invoices.Update_date, Tbl_Invoices.Document_Number, Tbl_Invoices.Date, Tbl_Invoices.Amount, Tbl_Invoices.Currency, Tbl_queries.Resolution_owner, Tbl_Invoices.mEMO, Tbl_Invoices.Overdue_Date, Tbl_queries.Query, Tbl_Customers.Credit_controller " & _
                                     "FROM Tbl_Areas RIGHT JOIN (Tbl_AdditionalQueryData RIGHT JOIN (Tbl_queries INNER JOIN (Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) ON Tbl_queries.ID = Tbl_Invoices.Query) ON (Tbl_AdditionalQueryData.Customer_Code = Tbl_Invoices.Customer_ID) AND (Tbl_AdditionalQueryData.Document_Number = Tbl_Invoices.Document_Number) AND (Tbl_AdditionalQueryData.Document_date = Tbl_Invoices.Date)) ON Tbl_Areas.ID = Tbl_Customers.Area " & _
                                     "WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yy") & "#) And (Tbl_Invoices.Amount>0) And (Tbl_queries.Query<>'Chargebacks') And ((Tbl_queries.Query) Is Not Null)) " & _
                                     "ORDER BY Tbl_Areas.Area, Tbl_Customers.Country, Tbl_Customers.Name, Tbl_Invoices.Date;")

    Rem Set rs = CurrentDb.OpenRecordset("SELECT Tbl_AdditionalQueryData.Query_date, Tbl_Areas.Area, Tbl_Customers.Country, Tbl_Customers.Name, Tbl_Customers.Customer_code, Tbl_Invoices.Update_date, Tbl_Invoices.Document_Number, Tbl_Invoices.Date, Tbl_Invoices.Amount, Tbl_Invoices.Currency, Tbl_queries.Resolution_owner, Tbl_Invoices.mEMO, Tbl_Invoices.Overdue_Date, Tbl_queries.Query, Tbl_Customers.Credit_controller " & _
    "FROM Tbl_Areas RIGHT JOIN (Tbl_AdditionalQueryData RIGHT JOIN (Tbl_queries INNER JOIN (Tbl_Customers INNER JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) ON Tbl_queries.ID = Tbl_Invoices.Query) ON (Tbl_AdditionalQueryData.Customer_Code = Tbl_Invoices.Customer_ID) AND (Tbl_AdditionalQueryData.Document_Number = Tbl_Invoices.Document_Number) AND (Tbl_AdditionalQueryData.Document_date = Tbl_Invoices.Date)) ON Tbl_Areas.ID = Tbl_Customers.Area " & _
    "WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yy") & "#) And (Tbl_Invoices.Amount>0) And ((Tbl_queries.Query) Is Not Null)) " & _
    "ORDER BY Tbl_Areas.Area, Tbl_Customers.Country, Tbl_Customers.Name, Tbl_Invoices.Date;")


    RS.MoveFirst
    If RS.RecordCount > 0 Then
        Set MyXl = CreateObject("excel.application")
        Set ExcDoc = MyXl.Workbooks.Open(GetPathExcelDirectory() & "\Retail Query Form.xls")
        MyXl.Visible = True
        ExcDoc.Worksheets.Add After:=ExcDoc.Sheets(3)

        With MyXl.ActiveSheet.PageSetup
            .LeftMargin = 0
            .RightMargin = 0
            .PrintQuality = 600
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 100
        End With

        ExcDoc.Worksheets(ExcDoc.Worksheets.Count).Name = Nz(RS.Fields("Area"), " ")
        row = 1
        With MyXl
            .Cells(1, 1) = DLookup("[Country]", "[Tbl_Countries]", "Code='" & RS.Fields("Country") & "'")
            .Cells(1, 2) = "Credit controller: " & GetNameCreditControllerFromID(RS.Fields("Credit_controller"))
            Call BoldLetter(MyXl, "A1:E1")
            Call FillCells(MyXl, "A1:E1", vbRed)
            row = 2
            Call QueryFileHeader(MyXl, row)
            row = 3
            FormulaRowStart = row
        End With
        While Not RS.EOF
            With MyXl
                CountryName = RS.Fields("Country")
                .Cells(row, 1) = Format(RS.Fields("Query_date"), "dd-mmm-yy")
                .Cells(row, 2) = RS.Fields("Country")
                .Cells(row, 3) = RS.Fields("Name")
                .Cells(row, 4) = RS.Fields("Customer_code")
                .Cells(row, 5) = RS.Fields("Document_Number")
                .Cells(row, 6) = Format(RS.Fields("Date"), "dd-mmm-yy")
                .Cells(row, 7) = Format(RS.Fields("Overdue_Date"), "dd-mmm-yy")
                .Cells(row, 8) = Format(RS.Fields("Amount"), "##,##0.00")
                .Cells(row, 9) = RS.Fields("Currency")
                .Cells(row, 10) = RS.Fields("Resolution_owner")
                .Cells(row, 11) = RS.Fields("Query")
                .Cells(row, 12) = RS.Fields("mEMO")
                row = row + 1
                RS.MoveNext
                If Not RS.EOF Then
                    If CountryName <> RS.Fields("Country") And _
                       ExcDoc.Worksheets(ExcDoc.Worksheets.Count).Name = RS.Fields("Area") Then
                        row = row + 1
                        .Cells(row, 7) = "Total:"
                        .Cells(row, 8) = "=Sum(H" & FormulaRowStart & ":H" & row - 1 & ")"
                        .Cells(row, 8).NumberFormat = "##,##0.00"
                        Call BoldLetter(MyXl, "G" & row & ":H" & row)
                        Call DoubleUnderlineCells(MyXl, "G" & row & ":H" & row)
                        row = row + 1
                        FormulaRowStart = row
                        .Cells(row, 1) = DLookup("[Country]", "[Tbl_Countries]", "Code='" & RS.Fields("Country") & "'")
                        .Cells(row, 2) = "Credit controller: " & GetNameCreditControllerFromID(RS.Fields("Credit_controller"))
                        Call BoldLetter(MyXl, "A" & row & ":E" & row)
                        Call FillCells(MyXl, "A" & row & ":E" & row, vbRed)

                        row = row + 1
                        Call QueryFileHeader(MyXl, row)
                        row = row + 1

                    End If
                End If

                If RS.EOF = False Then
                    If RS.Fields("Area") <> ExcDoc.Worksheets(ExcDoc.Worksheets.Count).Name Then
                        MyXl.Columns("A:M").AutoFit
                        MyXl.Columns("L").ColumnWidth = 30
                        MyXl.Columns("L").WrapText = True

                        row = row + 2
                        .Cells(row, 7) = "Total:"
                        .Cells(row, 8) = "=Sum(H" & FormulaRowStart & ":H" & row - 1 & ")"
                        Call BoldLetter(MyXl, "G" & row & ":H" & row)
                        Call DoubleUnderlineCells(MyXl, "G" & row & ":H" & row)
                        .Cells(row, 8).NumberFormat = "##,##0.00"
                        row = row + 1


                        ExcDoc.Worksheets.Add After:=ExcDoc.Sheets(ExcDoc.Worksheets.Count)
                        ExcDoc.Worksheets(ExcDoc.Worksheets.Count).Name = Nz(RS.Fields("Area"), " ")
                        row = 1
                        MyXl.Cells(1, 1) = DLookup("[Country]", "[Tbl_Countries]", "Code='" & RS.Fields("Country") & "'")
                        .Cells(1, 2) = "Credit controller: " & GetNameCreditControllerFromID(RS.Fields("Credit_controller"))
                        Call BoldLetter(MyXl, "A1:E1")
                        Call FillCells(MyXl, "A1:E1", vbRed)
                        row = 2
                        Call QueryFileHeader(MyXl, row)
                        row = 3
                        FormulaRowStart = row

                        With MyXl.ActiveSheet.PageSetup
                            .LeftMargin = 0
                            .RightMargin = 0
                            .PrintQuality = 600
                            .Orientation = xlLandscape
                            .PaperSize = xlPaperA4
                            .Zoom = False
                            .FitToPagesWide = 1
                            .FitToPagesTall = 100
                        End With

                    End If
                End If
            End With
        Wend
        With MyXl
            .Columns("A:M").AutoFit
            .Columns("L").ColumnWidth = 30
            .Columns("L").WrapText = True
            row = row + 2
            .Cells(row, 7) = "Total:"
            .Cells(row, 8) = "=Sum(H" & FormulaRowStart & ":H" & row - 1 & ")"
            Call BoldLetter(MyXl, "G" & row & ":H" & row)
            Call DoubleUnderlineCells(MyXl, "G" & row & ":H" & row)
            .Cells(row, 8).NumberFormat = "##,##0.00"
        End With
    End If

End Sub

Sub QueryFileHeader(ExcelFile As Variant, row As Integer)
    With ExcelFile
        .Cells(row, 1) = "Query Date"
        .Cells(row, 2) = "Country"
        .Cells(row, 3) = "Retail Customer"
        .Cells(row, 4) = "Bill To #"
        .Cells(row, 5) = "Invoice n#"
        .Cells(row, 6) = "Transaction Date"
        .Cells(row, 7) = "Due Date"
        .Cells(row, 8) = "Invoice Amt"
        .Cells(row, 9) = "Currency"
        .Cells(row, 10) = "ESDC/field office Owner"
        .Cells(row, 11) = "Query Type"
        .Cells(row, 12) = "Description"
        .Cells(row, 13) = "Next step"
        ExcelFile.Rows(row & ":" & row).RowHeight = 40.75
        Call BoldLetter(ExcelFile, "A" & row & ":O" & row)
        Call FillCells(ExcelFile, "A" & row & ":J" & row, vbYellow)
        Call FillCells(ExcelFile, "K" & row & ":M" & row, 10079487)

        Call BorderCells(ExcelFile, "A" & row & ":M" & row)
    End With
End Sub

Sub BoldLetter(ExcelFile As Variant, Coordinates As String)
    With ExcelFile
        .Range(Coordinates).Font.Bold = True
    End With
End Sub

Sub FillCells(ExcelFile As Variant, Coordinates As String, Colour As Long, Optional aThemeColor, Optional aTintAndShade As Variant)
    With ExcelFile
        If Colour <> 0 Then
            .Range(Coordinates).Interior.color = Colour
        Else
            .Range(Coordinates).Interior.ThemeColor = aThemeColor
            .Range(Coordinates).Interior.TintAndShade = aTintAndShade
        End If
    End With
End Sub

Sub DoubleUnderlineCells(ExcelFile As Variant, Coordinates As String)
    With ExcelFile
        With .Range(Coordinates).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
    End With
End Sub

Sub BorderCells(ExcelFile As Variant, Coordinates As String)
    With ExcelFile
        With .Range(Coordinates).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Range(Coordinates).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Range(Coordinates).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Range(Coordinates).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        With .Range(Coordinates).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

    End With
End Sub

Sub WriteOffsToPropose()
    DoCmd.OutputTo acOutputQuery, "QueryWriteOffs", acFormatXLS, QueryCLReport & GetPathExcelDirectory & "QueryWriteOffs.xls", True
End Sub

Sub CLLimitReport()
    Dim RS As Variant
    Dim MyXl As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row As Integer
    Dim FN As String

    Set RS = CurrentDb.OpenRecordset("SELECT Tbl_Areas.Area, Tbl_Countries.Country, Tbl_Customers.Name AS [Customer Name], Tbl_Customers.Customer_code AS [Customer ID], Tbl_CL.Currency AS [Currency code], Tbl_Customers.TotalInsurance AS [Insurance Credit Limit], Tbl_CL.CreditLimit AS [Logitech Credit Limit], [OpenARBalance]+[AwaitingInvoicing]+[AmtScheduledTom] AS [Exposure (with +3 day horizon)], [creditlimit]-([OpenARBalance]+[AwaitingInvoicing]+[AmtScheduledTom]) AS Available, Tbl_CL.OpenARBalance AS [Current O/S AR Balance], Tbl_CL.AwaitingInvoicing AS [Total Amt awaiting Invoicing], [CreditLimit]-[OpenARBalance]-[AwaitingInvoicing] AS [Shippable Balance], Tbl_CL.AmtScheduledTom AS [Total Amt Scheduled +3 days horizon] FROM ((Tbl_Customers LEFT JOIN Tbl_CL ON Tbl_Customers.Customer_code = Tbl_CL.Customer_code) LEFT JOIN Tbl_Countries ON Tbl_Customers.Country = Tbl_Countries.Code) LEFT JOIN Tbl_Areas ON Tbl_Customers.Area = Tbl_Areas.ID WHERE (((Tbl_CL.Currency) Is Not Null)) " & _
                                     " ORDER BY Tbl_Areas.Area, Tbl_Countries.Country, Tbl_Customers.Name;")

    Set MyXl = CreateObject("excel.application")
    Set ExcDoc = MyXl.Workbooks.Open(GetPathExcelDirectory() & "CL Report.xls")
    MyXl.ActiveSheet.Pictures.Insert (GetPathlogo())
    MyXl.ActiveSheet.Rows("1:1").RowHeight = 150
    row = 5
    With MyXl
        .Visible = False
        .Cells(row, 1) = "Area"
        .Cells(row, 2) = "Country"
        .Cells(row, 3) = "Customer name"
        .Cells(row, 4) = "Customer ID"
        .Cells(row, 5) = "Currency"
        .Cells(row, 6) = "Insurance CL"
        .Cells(row, 7) = "Logitech CL"
        .Cells(row, 8) = "Exposure (with +5 day horizon)"
        .Cells(row, 9) = "Available"
        .Cells(row, 10) = "Current O/S AR Balance"
        .Cells(row, 11) = "Total Amt awaiting Invoicing"
        .Cells(row, 12) = "Shippable Balance"
        .Cells(row, 13) = "Total Amt Scheduled +5 days horizon"

        .Rows("4:4").HorizontalAlignment = xlCenter
    End With

    row = 6
    With RS
        .MoveFirst
        While Not .EOF
            MyXl.Cells(row, 1) = RS.Fields(0)
            MyXl.Cells(row, 2) = RS.Fields(1)
            MyXl.Cells(row, 3) = RS.Fields(2)
            MyXl.Cells(row, 4) = RS.Fields(3)
            MyXl.Cells(row, 5) = RS.Fields(4)
            MyXl.Cells(row, 6) = Format(RS.Fields(5), "##,##0.00")
            MyXl.Cells(row, 7) = Format(RS.Fields(6), "##,##0.00")
            MyXl.Cells(row, 8) = Format(RS.Fields(7), "##,##0.00")
            MyXl.Cells(row, 9) = Format(RS.Fields(8), "##,##0.00")
            MyXl.Cells(row, 10) = Format(RS.Fields(9), "##,##0.00")
            MyXl.Cells(row, 11) = Format(RS.Fields(10), "##,##0.00")
            MyXl.Cells(row, 12) = Format(RS.Fields(11), "##,##0.00")
            MyXl.Cells(row, 13) = Format(RS.Fields(12), "##,##0.00")

            .MoveNext
            row = row + 1
        Wend
        With MyXl
            .Columns("A:M").AutoFit
            .Rows("5:5").AutoFilter

            .ActiveSheet.Shapes("Picture 1").Select
            .Selection.ShapeRange.IncrementLeft -5000
            .Selection.ShapeRange.IncrementLeft (MyXl.Sheets(1).Columns("A:M").Width - MyXl.ActiveSheet.Shapes("Picture 1").Width) / 2
            .Selection.ShapeRange.IncrementTop -5000
            .Selection.ShapeRange.IncrementTop 5

            .Cells(2, 1) = "Horizon Date Limit: " & Format(DLookup("CLHorizonDateLimit", "TblGeneral"), "dd-mmm-yyyy")
            .Cells(2, 1).Font.Bold = True
            .Cells(2, 1).Font.Italic = True
            .Cells(2, 1).Font.Size = 12

            .Cells(3, 1) = "Credit Limit Report printed on " & Format(Now, "dd-mmm-yyyy hh:mm:ss")
        End With

        With MyXl.ActiveSheet.PageSetup
            .PrintTitleRows = "$1:$4"
            .LeftMargin = 0
            .RightMargin = 0
            .PrintQuality = 600
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 100
        End With
        FN = GetPathExcelDirectory() & "Updated CL Report.xls"

        With MyXl
            .Application.DisplayAlerts = False
            .ActiveWorkbook.SaveAs FileName:=FN, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            .Application.DisplayAlerts = True
            .Quit
        End With

        Set MyXl = Nothing
    End With
End Sub

Sub OpenCLReport()
    Dim MyXl As Excel.Application
    Dim ExcDoc As Excel.Workbook

    Call CLLimitReport
    Set MyXl = CreateObject("excel.application")
    Set ExcDoc = MyXl.Workbooks.Open(GetPathExcelDirectory() & "Updated CL Report.xls")
    MyXl.Visible = True
End Sub

Sub UpdateHistoricalCL()
    Dim UpdateDate As Date
    Dim RefDate As Date
    Dim RS As Variant
    Dim RstHistoricalCLs As Variant

    Rem estrae il giorno lavorativo precedente
    UpdateDate = DateAdd("d", Date, -1)
    If Weekday(UpdateDate) = vbSunday Then
        UpdateDate = DateAdd("d", UpdateDate, -2)
    ElseIf Weekday(UpdateDate) = vbSaturday Then
        UpdateDate = DateAdd("d", UpdateDate, -1)
    End If


    Rem ##############
    Rem sostituire qui con la discriminante che stabilisce se e cosa cancellare
    If DCount("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd=#" & Format(UpdateDate, "mm/dd/yy") & "#") > 0 Then
        RefDate = DateAdd("d", -1, UpdateDate)
    Else
        RefDate = DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd<#" & Format(UpdateDate, "mm/dd/yy") & "#")
    End If
    CurrentDb.Execute ("DELETE Tbl_HistoricalCLsAndStatements.Update_date FROM Tbl_HistoricalCLsAndStatements WHERE (((Tbl_HistoricalCLsAndStatements.Update_date)>#" & Format(RefDate, "mm/dd/yy") & "#));")
    Rem ##############



    Rem inserisce in storico dati del credit limit relaviti al giorno lavorativo precedente
    CurrentDb.Execute _
        "INSERT INTO Tbl_HistoricalCLsAndStatements ( CreditLimit, OpenARBalance, AwaitingInvoicing, AmtScheduledTom, Update_date, Customer_code ) " & _
                                                                                                                                                     "SELECT Tbl_CL.CreditLimit, Tbl_CL.OpenARBalance, Tbl_CL.AwaitingInvoicing, Tbl_CL.AmtScheduledTom, #" & Format(UpdateDate, "mm/dd/yy") & "# AS Expr1, Tbl_CL.Customer_code " & _
                                                                                                                                                     "FROM Tbl_CL;"

    Rem inserisce Credit limit dell'assicurazione
    CurrentDb.Execute _
        "UPDATE Tbl_Customers INNER JOIN Tbl_HistoricalCLsAndStatements ON Tbl_Customers.Customer_code = Tbl_HistoricalCLsAndStatements.Customer_code SET Tbl_HistoricalCLsAndStatements.InsuranceCreditLimit = [Tbl_Customers].[TotalInsurance] " & _
                                                                                                                                                                                                                                                   "WHERE (((Tbl_HistoricalCLsAndStatements.Update_date)=#" & Format(UpdateDate, "mm/dd/yy") & "#));"


    If DCount("[Update_date]", "[Tbl_Invoices]", "Update_date=#" & Format(Date, "mm/dd/yy") & "#") > 0 Then
        Rem se e' stata gia' fatta l'importazione dgli estratti conto allora butta gli stessi dati in Tbl_HistoricalCLsAndStatements (trasfromando tutto in EUR)
        Set RS = CurrentDb.OpenRecordset("SELECT DISTINCT Tbl_Invoices.Customer_ID, Sum(IIf([Overdue_Date]>Date(),[Amount]*[exchangerate],0)) AS [current], Sum(IIf([Overdue_Date] Between Date() And DateAdd('d',Date(),-30),[Amount]*[exchangerate],0)) AS [1-30days], Sum(IIf([Overdue_Date] Between DateAdd('d',Date(),-31) And DateAdd('d',Date(),-60),[Amount]*[exchangerate],0)) AS [31-60days], Sum(IIf([Overdue_Date] Between DateAdd('d',Date(),-61) And DateAdd('d',Date(),-90),[Amount]*[exchangerate],0)) AS [61-90days], Sum(IIf([Overdue_Date]<DateAdd('d',Date(),-90),[Amount]*[exchangerate],0)) AS Over90days " & _
                                         "FROM Tbl_Invoices INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yy") & "#)) GROUP BY Tbl_Invoices.Customer_ID;")
        RS.MoveFirst
        Set RstHistoricalCLs = CurrentDb.OpenRecordset("SELECT Tbl_HistoricalCLsAndStatements.* FROM Tbl_HistoricalCLsAndStatements;")
        While RS.EOF = False
            RstHistoricalCLs.FindFirst "Customer_code=" & RS.Fields("Customer_ID") & " AND Update_date=#" & Format(UpdateDate, "mm/dd/yy") & "#"
            If RstHistoricalCLs.NoMatch = False Then
                RstHistoricalCLs.Edit
                RstHistoricalCLs.Fields("Current") = RS.Fields("Current")
                RstHistoricalCLs.Fields("Overdue1-30Days") = RS.Fields("1-30days")
                RstHistoricalCLs.Fields("Overdue31-60Days") = RS.Fields("31-60days")
                RstHistoricalCLs.Fields("Overdue61-90Days") = RS.Fields("61-90days")
                RstHistoricalCLs.Fields("OverdueOver90Days") = RS.Fields("Over90days")
                RstHistoricalCLs.Update
            End If
            RS.MoveNext
        Wend
        RstHistoricalCLs.Close
        Set RstHistoricalCLs = Nothing
    End If

End Sub

Sub WAVPlay(file)
    Dim SoundName As String
End Sub

Function RemoveNotNumericChar(Stringa As String) As String
    Dim Counter As Long
    Counter = 1
    Do While Counter <= Len(Stringa)
        If Not (Asc(Mid(Stringa, Counter, 1)) >= Asc(0) And Asc(Mid(Stringa, Counter, 1)) <= Asc(9)) Then
            Stringa = Replace(Stringa, Mid(Stringa, Counter, 1), "")
        Else
            Counter = Counter + 1
        End If
    Loop
    RemoveNotNumericChar = Stringa
End Function

Sub OpenHelpPage(MskName As String)
    Dim strURL_c As String
    Dim objIE As SHDocVw.InternetExplorer
    Dim ieDoc As MSHTML.HTMLDocument
    Dim tbxPwdFld As MSHTML.HTMLInputElement
    Dim tbxUsrFld As MSHTML.HTMLInputElement
    Dim btnSubmit As MSHTML.HTMLInputElement
    On Error GoTo Err_Hnd
    'Create Internet Explorer Object
    Set objIE = New SHDocVw.InternetExplorer
    'strURL_c = "http://exchange.logitech.com/docs/DOC-10483"
    strURL_c = Nz(DLookup("[HelpPage]", "[Tbl_HelpPages]", "MskName='" & MskName & "'"), "https://exchange.logitech.com/docs/DOC-10614")
    'Navigate the URL
    objIE.Navigate strURL_c
    'Wait for page to load
    Do Until objIE.ReadyState = READYSTATE_COMPLETE: Loop
    'Get document object
    Set ieDoc = objIE.Document
    'Get username/password fields and submit button.
    '        Set tbxPwdFld = ieDoc.all.Item("Passwd")
    '       Set tbxUsrFld = ieDoc.all.Item("Email")
    '      Set btnSubmit = ieDoc.all.Item("signIn")
    'Fill Fields
    tbxUsrFld.value = ""
    tbxPwdFld.value = ""
    'Click submit
    'btnSubmit.Click
    'Wait for page to load
    '    Application.FollowHyperlink Nz(DLookup("[HelpPage]", "[Tbl_HelpPages]", "MskName='" & MskName & "'"), 0)
    '"http://exchange.logitech.com/docs/DOC-10483"

    Do Until objIE.ReadyState = READYSTATE_COMPLETE: Loop
Err_Hnd:                                         '(Fail gracefully)
    objIE.Visible = True
End Sub

Function FillGeneralCashTargetByEmail() As String
    Dim qdfNew As DAO.QueryDef
    Dim SQLText As String
    Dim rst2 As Variant
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row, col, I, TotalQuarter As Integer
    Dim CashTargetDate As Date
    'Dim UDSExchangeRate As Currency
    FillGeneralCashTargetByEmail = ""


    With CurrentDb

        '   UDSExchangeRate = DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='USD'")

        Set rst2 = CurrentDb.OpenRecordset("SELECT TOP 1 Tbl_MonthEnd.LatestPaymentDate, Tbl_MonthEnd.FiscalYear, Tbl_MonthEnd.FiscalQuarter, Tbl_MonthEnd.FiscalMonth FROM Tbl_MonthEnd WHERE (((Tbl_MonthEnd.MonthEnd)>=#" & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "mm/dd/yy") & "#)) ;")

        SQLText = "SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, Round(Sum([Tbl_CashCollected].[amount]/1000)) AS TotalAmount FROM Tbl_CashCollected GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
                  "UNION SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000)) AS TotalAmount FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000)) HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"

        '    SqlText = "SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, (Round(Sum([Tbl_CashCollected].[amount]/1000))/" & UDSExchangeRate & ") AS TotalAmount FROM Tbl_CashCollected GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
        '             "UNION SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round((([Tbl_Cash_Target_Breakdown].[Amount] / [Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency] / 1000)) / " & UDSExchangeRate & ") As TotalAmount FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round ((([Tbl_Cash_Target_Breakdown].[Amount] / [Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency] / 1000)) / " & UDSExchangeRate & ") HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"

        Set qdfNew = .CreateQueryDef("Query1", SQLText)
        DoCmd.OutputTo acOutputQuery, "Query1", acFormatXLS, GetPathExcelDirectory() & "GeneralCashTarget.xls", False
        .QueryDefs.Delete ("Query1")

        Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
        Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "GeneralCashTarget.xls")
        ExcApp.Visible = False
        Rem ExcApp.visible = True
        With ExcDoc
            row = 1
            While .ActiveSheet.Cells(row + 1, 1) <> ""
                row = row + 1
            Wend
            .Sheets.Add
            ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                                     "Query1!R1C1:R" & row & "C6", Version:=xlPivotTableVersion10).CreatePivotTable _
                                                     TableDestination:="Sheet1!R3C1", tableName:="PivotTable1", DefaultVersion _
                                                     :=xlPivotTableVersion10
            .Sheets("Sheet1").Select
            row = 3
            col = 1
            With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("Description")
                .Orientation = xlRowField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable1").AddDataField .ActiveSheet.PivotTables( _
                                                                 "PivotTable1").PivotFields("TotalAmount"), "Sum of TotalAmount", xlSum
            With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("StartDate")
                .Orientation = xlColumnField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable1").PivotFields("Description").AutoSort _
        xlDescending, "Description"
            .ActiveSheet.Range("B5:E6").NumberFormat = "#,##0"
            .ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
            .ActiveSheet.Range("A1:IV100").Copy
            .ActiveSheet.Range("A1:IV100").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                        :=False, Transpose:=False

            For I = 1 To 3
                If .ActiveSheet.Cells(row + 1, I + 1) <> "" Then
                    If IsDate(.ActiveSheet.Cells(row + 1, I + 1)) Then
                        ' .ActiveSheet.Cells(Row + 1, i + 1) = DateAdd("yyyy", -1, .ActiveSheet.Cells(Row + 1, i + 1))
                    End If
                    .ActiveSheet.Cells(row + 1, I + 1).NumberFormat = """M" & I & """ mmm yy"
                End If
            Next I

            With .ActiveSheet.Range("A1:IV100")
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With
            .ActiveSheet.Cells(row, col) = "Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(row + 1, col) = DLookup("Area", "tblGeneral") & " TOTAL                               " & "(" & DLookup("MainCurrency", "tblGeneral") & ")"

            .ActiveSheet.Cells(row, col + 1) = ""
            .ActiveSheet.Range("b" & row + 1 & ":d" & row + 1).Copy
            .ActiveSheet.Range("b" & row & ":d" & row).PasteSpecial
            .ActiveSheet.Cells(row + 1, col + 1) = ""
            .ActiveSheet.Cells(row + 1, col + 2) = ""
            .ActiveSheet.Cells(row + 1, col + 3) = ""
            .ActiveSheet.Cells(row + 1, col + 4) = ""
            I = 1
            While .ActiveSheet.Cells(row, I) <> ""
                I = I + 1
            Wend
            '        i = i - 1
            .ActiveSheet.Cells(row, I) = "TOTAL Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            TotalQuarter = I
            Call BoldLetter(ExcApp, "A1:Z4")
            row = row + 4
            .ActiveSheet.Cells(row, 1) = "Actual Performance to date"
            For I = 1 To 4
                If .ActiveSheet.Cells(row - 1, col + I) <> "" Then
                    If (.ActiveSheet.Range(Chr(65 + I) & row - 2) = "") Or (.ActiveSheet.Range(Chr(65 + I) & row - 2) = 0) Then
                        .ActiveSheet.Cells(row, col + I) = "-"
                        .ActiveSheet.Cells(row, col + I).HorizontalAlignment = xlRight
                    Else
                        .ActiveSheet.Cells(row, col + I) = "=" & Chr(65 + I) & row - 1 & "/" & Chr(65 + I) & row - 2
                    End If
                    .ActiveSheet.Cells(row, col + I).NumberFormat = "0%"
                End If
            Next I
            Call BoldLetter(ExcApp, "A" & row & ":Z" & row)
            .ActiveSheet.Range("A4..Z4").Insert Shift:=xlDown
            .ActiveSheet.Range("A6..Z6").Insert Shift:=xlDown
            Call FillCells(ExcApp, "A5:A5", 0, xlThemeColorAccent5, 0.599993896298105)
            I = 1
            While .ActiveSheet.Cells(3, I + 1) <> ""
                I = I + 1
            Wend
            Call FillCells(ExcApp, "A2:" & Chr(64 + I) & "4", 5296274)
            ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        End With
        Call SetFontColor(ExcApp, "A7:E7", -4165632)
        Call SetFontColor(ExcApp, "A8:E9", -6279056)
        Call BoldLetter(ExcApp, "A4:E40")
        ExcDoc.ActiveSheet.Cells(3, 1).HorizontalAlignment = xlCenter
        Call BorderCells(ExcApp, "A" & 2 & ":" & Chr(64 + I) & row + 2)
        Call BorderCells(ExcApp, "A5:" & Chr(64 + I) & row + 2)

        ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        ExcDoc.Worksheets("Sheet1").Cells(1, 1) = "GENERAL CASH TARGET"
        '    CashTargetDate = DateAdd("d", -1, Date)
        '   While Weekday(CashTargetDate) = vbSaturday Or Weekday(CashTargetDate) = vbSunday
        '      CashTargetDate = DateAdd("d", -1, CashTargetDate)
        ' Wend
        CashTargetDate = Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "DD MMM YYyy")
        ExcDoc.Worksheets("Sheet1").Cells(1, 1) = "GENERAL CASH TARGET AS OF " & Format(CashTargetDate, "dd mmmm yyyy")
        ExcDoc.ActiveSheet.Range("A1:" & Chr(64 + TotalQuarter) & "1").HorizontalAlignment = xlCenterAcrossSelection

        With ExcApp
            .Application.DisplayAlerts = False
            .ActiveWorkbook.SaveAs FileName:=GetPathExcelDirectory() & "GeneralCashTarget.xls", FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            .Application.DisplayAlerts = True
            .Quit
        End With

        Set ExcDoc = Nothing
        Set ExcApp = Nothing
        FillGeneralCashTargetByEmail = GetPathExcelDirectory() & "GeneralCashTarget.xls"
    End With

End Function

Function FillCashTargetWithChannelByEmail()
    Dim qdfNew As DAO.QueryDef
    Dim SQLText, StringCurrency As String
    Dim rst2 As Variant
    Dim rst As Variant
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row, col, I, NumQuarters, StartingMonth As Integer
    Dim ValExchangeRate As Currency

    With CurrentDb
        Set rst2 = CurrentDb.OpenRecordset("SELECT TOP 1 Tbl_MonthEnd.LatestPaymentDate, Tbl_MonthEnd.FiscalYear, Tbl_MonthEnd.FiscalQuarter, Tbl_MonthEnd.FiscalMonth FROM Tbl_MonthEnd WHERE (((Tbl_MonthEnd.MonthEnd)>=#" & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "mm/dd/yy") & "#)) ;")


        SQLText = "SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, round((([Tbl_Cash_Target_Breakdown].[amount] / [Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency])/1000)) AS TotalAmount, Tbl_Cash_Target_Breakdown.Channel FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000)) , Tbl_Cash_Target_Breakdown.Channel HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
                  "UNION SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, Round(Sum([Tbl_CashCollected].[amount]/1000)) AS TotalAmount, Tbl_Customers.RetailOEM AS Channel FROM Tbl_CashCollected LEFT JOIN Tbl_Customers ON Tbl_CashCollected.CustomerID = Tbl_Customers.Customer_code " & _
                  "GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]), Tbl_Customers.RetailOEM HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"

        'Round(([Tbl_Cash_Target_Breakdown].[amount]*[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000))

        Set qdfNew = .CreateQueryDef("Query1", SQLText)
        DoCmd.OutputTo acOutputQuery, "Query1", acFormatXLS, GetPathExcelDirectory() & "GeneralCashTargetWithChannel.xls"
        .QueryDefs.Delete ("Query1")

        Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
        Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "GeneralCashTargetWithChannel.xls")
        ExcApp.Visible = False
        Rem ExcApp.visible = True

        Call FixCashReport(ExcDoc)

        With ExcDoc
            row = 1
            While .ActiveSheet.Cells(row, 1) <> ""
                row = row + 1
            Wend
            row = row - 1
            .Sheets.Add

            ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                                     "Query1!R1C1:R" & row & "C7", Version:=xlPivotTableVersion10).CreatePivotTable _
                                                     TableDestination:="Sheet1!R3C1", tableName:="PivotTable2", DefaultVersion _
                                                     :=xlPivotTableVersion10
            .Sheets("Sheet1").Select
            .ActiveSheet.Cells(3, 1).Select
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("Channel")
                .Orientation = xlRowField
                .Position = 1
            End With
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description")
                .Orientation = xlRowField
                .Position = 2
            End With
            .ActiveSheet.PivotTables("PivotTable2").AddDataField .ActiveSheet.PivotTables( _
                                                                 "PivotTable2").PivotFields("TotalAmount"), "Sum of TotalAmount", xlSum
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("StartDate")
                .Orientation = xlColumnField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description").AutoSort _
        xlDescending, "Description"
            .ActiveSheet.Range("A4").Select
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("Channel").Subtotals = Array _
                                                                                       (False, False, False, False, False, False, False, False, False, False, False, False)
            .ActiveSheet.PivotTables("PivotTable2").PivotSelect "Description", xlButton, _
                                                                True
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description").Subtotals = _
                                                                                         Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .ActiveSheet.Range("B8").Select
            .ActiveSheet.PivotTables("PivotTable2").InGridDropZones = False

            .ActiveSheet.Range("A1:IV100").Copy
            .ActiveSheet.Range("A1:IV100").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                        :=False, Transpose:=False

            NumQuarters = 0
            For I = 1 To 3
                If IsDate(.ActiveSheet.Cells(4, I + 2)) Then
                    '                .ActiveSheet.Cells(4, i + 2) = DateAdd("yyyy", -1, .ActiveSheet.Cells(4, i + 2))
                    .ActiveSheet.Cells(4, I + 2).NumberFormat = """M" & I & """ mmm yy"
                    .ActiveSheet.Cells(4, I + 2).HorizontalAlignment = xlCenter
                    col = I + 2
                    NumQuarters = NumQuarters + 1
                End If
            Next I

            .ActiveSheet.Cells(3, 1) = ""
            .ActiveSheet.Cells(4, 2) = "Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(4, 2).HorizontalAlignment = xlCenter

            .ActiveSheet.Cells(4, NumQuarters + 3) = "TOTAL Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(4, NumQuarters + 3).HorizontalAlignment = xlCenter
            Call BoldLetter(ExcApp, "A4:Z4")
            .ActiveSheet.Cells(3, 3) = ""
            .ActiveSheet.Cells(4, 1) = ""

            .ActiveSheet.Range("C5:F60").NumberFormat = "#,##0"

            .ActiveSheet.Range("A5..Z5").Insert Shift:=xlDown

            With .ActiveSheet.Range("A1:IV100")
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With

            Call BorderCells(ExcApp, "B3:" & Chr(66 + NumQuarters + 1) & 5)

            row = 6
            While .ActiveSheet.Cells(row, 2) <> ""
                If .ActiveSheet.Cells(row, 2) <> "" Then
                    StringCurrency = Nz(DLookup("ReportCurrency", "tbl_channels", "Name='" & .ActiveSheet.Cells(row, 1) & "'"), DLookup("MainCurrency", "TblGeneral"))
                    '               ValExchangeRate = DLookup("ExchangeRate", "tbl_currencies", "CurrencyID='" & StringCurrency & "'")
                    StartingMonth = (DMin("FiscalMonth", "Tbl_Cash_Target_Breakdown", "FiscalYear=" & rst2.Fields("FiscalYear") & " AND FiscalQuarter=" & rst2.Fields("FiscalQuarter"))) - 1
                    For I = 1 To NumQuarters
                        '       Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.OriginalCurrency FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.OriginalCurrency HAVING (((Tbl_Cash_Target_Breakdown.OriginalCurrency)=" & StringCurrency & ") AND ((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalMonth)=" & StartingMonth + 1 & "));")
                        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth HAVING (((Tbl_Cash_Target_Breakdown.OriginalCurrency)='" & StringCurrency & "') AND ((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalMonth)=" & StartingMonth + I & ")); ")

                        .ActiveSheet.Cells(row, 2 + I) = "=" & CCur(.ActiveSheet.Cells(row, 2 + I)) * rst.Fields("ExchangeRateToMainCurrency")
                        .ActiveSheet.Cells(row, 2 + NumQuarters + 1) = "=Sum(C" & row & ":" & Chr(67 + NumQuarters - 1) & row & ")"
                        .ActiveSheet.Cells(row + 1, 2 + I) = "=" & CCur(.ActiveSheet.Cells(row + 1, 2 + I)) * rst.Fields("ExchangeRateToMainCurrency")
                        .ActiveSheet.Cells(row + 1, 2 + NumQuarters + 1) = "=Sum(C" & row + 1 & ":" & Chr(67 + NumQuarters - 1) & row + 1 & ")"
                    Next I
                    '.ActiveSheet.Cells(Row, 2 + NumQuarters + 1) = "=" & CCur(.ActiveSheet.Cells(Row, 2 + NumQuarters + 1)) / rst.Fields("ExchangeRateToMainCurrency")

                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown
                    .ActiveSheet.Cells(row, 2) = Left(.ActiveSheet.Cells(row + 1, 1) & "                                                                       ", 48) & "(" & StringCurrency & ")"
                    Call SetFontColor(ExcApp, "B" & row & ":E" & row, vbBlack)
                    .ActiveSheet.Range("B" & row & ":B" & row).HorizontalAlignment = xlLeft
                    row = row + 1
                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown

                    Call BoldLetter(ExcApp, "B" & row - 1 & ":B" & row - 1)
                    Call FillCells(ExcApp, "B" & row - 1 & ":B" & row - 1, 0, xlThemeColorAccent5, 0.599993896298105)
                    .ActiveSheet.Range("B" & row & ":B" & row).HorizontalAlignment = xlLeft
                    .ActiveSheet.Cells(row + 1, 1) = ""
                    row = row + 1
                    Call SetFontColor(ExcApp, "A" & row & ":F" & row, -4165632)
                    Call BoldLetter(ExcApp, "A" & row & ":Z" & row)
                    Call SetFontColor(ExcApp, "A" & row + 1 & ":F" & row + 1, -6279056)
                    Call BoldLetter(ExcApp, "A" & row + 1 & ":Z" & row + 1)

                    row = row + 2
                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown
                    .ActiveSheet.Cells(row, 2) = "Actual Performance to date"





                    For I = 1 To NumQuarters + 1
                        If (.ActiveSheet.Cells(row - 2, 2 + I) = "") Or (.ActiveSheet.Cells(row - 2, 2 + I) = 0) Then
                            .ActiveSheet.Cells(row, 2 + I) = "-"
                            .ActiveSheet.Cells(row, 2 + I).HorizontalAlignment = xlRight
                        Else
                            .ActiveSheet.Cells(row, 2 + I) = "=" & (Chr(66 + I)) & row - 1 & "/" & (Chr(66 + I)) & row - 2
                        End If
                        .ActiveSheet.Cells(row, 2 + I).NumberFormat = "0%"

                    Next I


                    Rem
                    Rem             For I = 1 To NumQuarters + 1
                    Rem              .ActiveSheet.Cells(Row, 2 + I) = "=" & (Chr(66 + I)) & Row - 1 & "/" & (Chr(66 + I)) & Row - 2
                    Rem           .ActiveSheet.Cells(Row, 2 + I).NumberFormat = "0%"
                    Rem    Next I







                    Call BorderCells(ExcApp, "B" & row - 4 & ":" & Chr(66 + NumQuarters + 1) & row)
                    Call SetFontColor(ExcApp, "B" & row & ":G" & row, -6279056)
                    Call BoldLetter(ExcApp, "B" & row & ":G" & row)
                End If
                row = row + 1
            Wend
            Call FillCells(ExcApp, "B3:" & Chr(64 + col + 1) & "5", 5296274)
        End With
        ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        ExcDoc.Worksheets("Sheet1").Cells(1, 2) = "BREAKDOWN BY CHANNEL"
        Call BoldLetter(ExcApp, "B1:B1")
        ExcDoc.ActiveSheet.Range("B1:" & Chr(66 + 1 + NumQuarters) & "1").HorizontalAlignment = xlCenterAcrossSelection

    End With

    With ExcApp
        .Application.DisplayAlerts = False
        .ActiveWorkbook.SaveAs FileName:=GetPathExcelDirectory() & "GeneralCashTargetWithChannel.xls", FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        .Application.DisplayAlerts = True
        .Quit
    End With

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    FillCashTargetWithChannelByEmail = GetPathExcelDirectory() & "GeneralCashTargetWithChannel.xls"
End Function

Function FillCashTargetWithCurrencyByEmail()
    Dim qdfNew As DAO.QueryDef
    Dim SQLText, StringCurrency As String
    Dim rst2 As Variant
    Dim rst As Variant
    Dim rst3 As Variant
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row, col, I, NumQuarters, StartingMonth As Integer
    Dim ValExchangeRate As Currency

    With CurrentDb
        Set rst2 = CurrentDb.OpenRecordset("SELECT TOP 1 Tbl_MonthEnd.LatestPaymentDate, Tbl_MonthEnd.FiscalYear, Tbl_MonthEnd.FiscalQuarter, Tbl_MonthEnd.FiscalMonth FROM Tbl_MonthEnd WHERE (((Tbl_MonthEnd.MonthEnd)>=#" & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "mm/dd/yy") & "#)) ;")

        SQLText = "SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000))  AS TotalAmount, Tbl_Cash_Target_Breakdown.OriginalCurrency FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Tbl_Cash_Target_Breakdown.OriginalCurrency, Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency]/1000)) HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));" & _
                  "UNION SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, Round(Sum([Tbl_CashCollected].[amount]/1000)) AS TotalAmount, Tbl_CashCollected.Currency AS OriginalCurrency FROM Tbl_CashCollected LEFT JOIN Tbl_Customers ON Tbl_CashCollected.CustomerID = Tbl_Customers.Customer_code GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]), Tbl_CashCollected.Currency HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"


        Set qdfNew = .CreateQueryDef("Query1", SQLText)
        DoCmd.OutputTo acOutputQuery, "Query1", acFormatXLS, GetPathExcelDirectory() & "GeneralCashTargetWithCurrency.xls"
        .QueryDefs.Delete ("Query1")

        Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
        Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "GeneralCashTargetWithCurrency.xls")
        ExcApp.Visible = False
        Rem ExcApp.visible = True

        Call FixCashReport(ExcDoc)

        With ExcDoc
            row = 1
            While .ActiveSheet.Cells(row, 1) <> ""
                row = row + 1
            Wend
            row = row - 1

            .Sheets.Add
            ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                                     "Query1!R1C1:R" & row & "C7", Version:=xlPivotTableVersion10).CreatePivotTable _
                                                     TableDestination:="Sheet1!R3C1", tableName:="PivotTable2", DefaultVersion _
                                                     :=xlPivotTableVersion10
            .Sheets("Sheet1").Select
            .ActiveSheet.PivotTables("PivotTable2").AddDataField .ActiveSheet.PivotTables( _
                                                                 "PivotTable2").PivotFields("TotalAmount"), "Sum of TotalAmount", xlSum
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("StartDate")
                .Orientation = xlColumnField
                .Position = 1
            End With
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("OriginalCurrency")
                .Orientation = xlRowField
                .Position = 1
            End With
            With .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description")
                .Orientation = xlRowField
                .Position = 2
            End With
            .ActiveSheet.PivotTables("PivotTable2").PivotSelect "Description[All]", _
                                                                xlLabelOnly, True
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description").AutoSort _
        xlDescending, "Description"
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("Description").Subtotals = _
                                                                                         Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .ActiveSheet.PivotTables("PivotTable2").PivotFields("OriginalCurrency"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
                          False, False)

            .ActiveSheet.Range("A1:IV100").Copy
            .ActiveSheet.Range("A1:IV100").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                        :=False, Transpose:=False

            NumQuarters = 0

            For I = 1 To 3
                If IsDate(.ActiveSheet.Cells(4, I + 2)) Then
                    '.ActiveSheet.Cells(4, i + 2) = DateAdd("yyyy", -1, .ActiveSheet.Cells(4, i + 2))
                    .ActiveSheet.Cells(4, I + 2).NumberFormat = """M" & I & """ mmm yy"
                    .ActiveSheet.Cells(4, I + 2).HorizontalAlignment = xlCenter
                    col = I + 2
                    NumQuarters = NumQuarters + 1
                End If
            Next I

            .ActiveSheet.Cells(3, 1) = ""
            .ActiveSheet.Cells(4, 2) = "Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(4, 2).HorizontalAlignment = xlCenter

            .ActiveSheet.Cells(4, 2 + NumQuarters + 1) = "TOTAL Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(4, 2 + NumQuarters + 1).HorizontalAlignment = xlCenter
            Call BoldLetter(ExcApp, "A4:Z4")
            .ActiveSheet.Cells(3, 3) = ""
            .ActiveSheet.Cells(4, 1) = ""

            .ActiveSheet.Range("C5:F60").NumberFormat = "#,##0"

            .ActiveSheet.Range("A5..Z5").Insert Shift:=xlDown
            With .ActiveSheet.Range("A1:IV100")
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With

            Call BorderCells(ExcApp, "B3:" & Chr(66 + NumQuarters + 1) & 5)

            row = 6

            While .ActiveSheet.Cells(row, 2) <> ""
                If .ActiveSheet.Cells(row, 2) <> "" Then
                    StringCurrency = .ActiveSheet.Cells(row, 1)
                    Set rst3 = CurrentDb.OpenRecordset("SELECT Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.Channel FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.Channel HAVING (((Tbl_Cash_Target_Breakdown.OriginalCurrency)='" & StringCurrency & "'));")
                    StartingMonth = (DMin("FiscalMonth", "Tbl_Cash_Target_Breakdown", "FiscalYear=" & rst2.Fields("FiscalYear") & " AND FiscalQuarter=" & rst2.Fields("FiscalQuarter"))) - 1
                    For I = 1 To NumQuarters
                        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMonth HAVING (((Tbl_Cash_Target_Breakdown.OriginalCurrency)='" & StringCurrency & "') AND ((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalMonth)=" & StartingMonth + I & ")); ")

                        .ActiveSheet.Cells(row, 2 + I) = "=" & CCur(.ActiveSheet.Cells(row, 2 + I)) * rst.Fields("ExchangeRateToMainCurrency")
                        .ActiveSheet.Cells(row, 2 + NumQuarters + 1) = "=Sum(C" & row & ":" & Chr(67 + NumQuarters - 1) & row & ")"
                        .ActiveSheet.Cells(row + 1, 2 + I) = "=" & CCur(.ActiveSheet.Cells(row + 1, 2 + I)) * rst.Fields("ExchangeRateToMainCurrency")
                        .ActiveSheet.Cells(row + 1, 2 + NumQuarters + 1) = "=Sum(C" & row + 1 & ":" & Chr(67 + NumQuarters - 1) & row + 1 & ")"

                    Next I
                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown

                    StringCurrency = ""
                    rst3.MoveFirst
                    While Not rst3.EOF
                        StringCurrency = StringCurrency & rst3.Fields("Channel") & ", "
                        rst3.MoveNext
                    Wend
                    StringCurrency = Left(StringCurrency, Len(StringCurrency) - 2)
                    .ActiveSheet.Cells(row, 2) = UCase(ExcDoc.ActiveSheet.Cells(row + 1, 1) & String(60 - Len(StringCurrency) - Len(.ActiveSheet.Cells(row + 1, 1)), " ") & "(" & StringCurrency & ")")
                    Call SetFontColor(ExcApp, "B" & row & ":E" & row, vbBlack)
                    .ActiveSheet.Range("B" & row & ":B" & row).HorizontalAlignment = xlLeft
                    row = row + 1
                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown

                    Call BoldLetter(ExcApp, "B" & row - 1 & ":B" & row - 1)
                    Call FillCells(ExcApp, "B" & row - 1 & ":B" & row - 1, 0, xlThemeColorAccent5, 0.599993896298105)
                    .ActiveSheet.Range("B" & row & ":B" & row).HorizontalAlignment = xlLeft
                    .ActiveSheet.Cells(row + 1, 1) = ""
                    row = row + 1
                    Call SetFontColor(ExcApp, "A" & row & ":F" & row, -4165632)
                    Call BoldLetter(ExcApp, "A" & row & ":Z" & row)
                    Call SetFontColor(ExcApp, "A" & row + 1 & ":F" & row + 1, -6279056)
                    Call BoldLetter(ExcApp, "A" & row + 1 & ":Z" & row + 1)

                    row = row + 2
                    .ActiveSheet.Range("A" & row & "..Z" & row).Insert Shift:=xlDown
                    .ActiveSheet.Cells(row, 2) = "Actual Performance to date"

                    For I = 1 To NumQuarters + 1
                        If (.ActiveSheet.Cells(row - 2, 2 + I) = "") Or (.ActiveSheet.Cells(row - 2, 2 + I) = 0) Then
                            .ActiveSheet.Cells(row, 2 + I) = "-"
                            .ActiveSheet.Cells(row, 2 + I).HorizontalAlignment = xlRight
                        Else
                            .ActiveSheet.Cells(row, 2 + I) = "=" & (Chr(66 + I)) & row - 1 & "/" & (Chr(66 + I)) & row - 2
                        End If
                        .ActiveSheet.Cells(row, 2 + I).NumberFormat = "0%"

                    Next I








                    Rem               For I = 1 To NumQuarters + 1
                    Rem                .ActiveSheet.Cells(Row, 2 + I) = "=" & (Chr(66 + I)) & Row - 1 & "/" & (Chr(66 + I)) & Row - 2
                    Rem             .ActiveSheet.Cells(Row, 2 + I).NumberFormat = "0%"
                    Rem      Next I
                    Call BorderCells(ExcApp, "B" & row - 4 & ":" & Chr(66 + NumQuarters + 1) & row)
                    Call SetFontColor(ExcApp, "B" & row & ":G" & row, -6279056)
                    Call BoldLetter(ExcApp, "B" & row & ":G" & row)
                End If
                row = row + 1
            Wend
            Call FillCells(ExcApp, "B3:" & Chr(64 + col + 1) & "5", 5296274)
        End With
        ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        ExcDoc.Worksheets("Sheet1").Cells(1, 2) = "BREAKDOWN BY CURRENCY"
        Call BoldLetter(ExcApp, "B1:B1")
        ExcDoc.ActiveSheet.Range("B1:" & Chr(66 + 1 + NumQuarters) & "1").HorizontalAlignment = xlCenterAcrossSelection

    End With

    With ExcApp
        .Application.DisplayAlerts = False
        .ActiveWorkbook.SaveAs FileName:=GetPathExcelDirectory() & "GeneralCashTargetWithCurrency.xls", FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
        .Application.DisplayAlerts = True
        .Quit
    End With

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    FillCashTargetWithCurrencyByEmail = GetPathExcelDirectory() & "GeneralCashTargetWithCurrency.xls"

End Function

Function MergeAllCashTargetRerports(MainCashTargetReport, CashTargetWithChannelReport, CashTargetWithChannelByCurrencyReport, MainCashTargetReportInUSD)
    Dim ExcApp, ExcApp2 As Excel.Application
    Dim ExcDoc, ExcDoc2 As Excel.Workbook
    Dim I, TopNextReport, BottomNextReport As Integer
    Dim ReportName As String

    Dim vFile As Variant
    Dim sFilter As String, lPicType As Long
    Dim oPic As IPictureDisp

    TopNextReport = 1
    BottomNextReport = 10
    Set ExcApp = CreateObject("Excel.Application") 'apre il nuovo modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Add
    ExcApp.Visible = False

    Set ExcApp2 = CreateObject("Excel.Application") 'apre il primo prospetto
    Set ExcDoc2 = ExcApp2.Workbooks.Open(MainCashTargetReport)
    ExcApp2.Visible = False

    With ExcDoc
        ExcDoc2.ActiveSheet.Range("A" & TopNextReport & ":O" & BottomNextReport).Copy
        .ActiveSheet.Range("A" & TopNextReport & ":O" & BottomNextReport).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        TopNextReport = BottomNextReport + 1
        ExcApp2.CutCopyMode = False
        ExcDoc2.Close
        Set ExcDoc2 = Nothing

        Set ExcApp2 = CreateObject("Excel.Application") 'apre il secondo prospetto
        Set ExcDoc2 = ExcApp2.Workbooks.Open(CashTargetWithChannelReport)
        ExcApp2.Visible = False
        I = 1
        While ExcDoc2.ActiveSheet.Cells(I, 1) = ""
            I = I + 1
        Wend
        BottomNextReport = I - 1
        I = 1
        While InStr(1, ExcDoc2.ActiveSheet.Cells(4, I), "TOTAL Q") = 0
            I = I + 1
        Wend
        I = I - 1
        ExcDoc2.ActiveSheet.Range("B1:" & Chr(Asc("A") + I) & BottomNextReport).Copy
        '.ActiveSheet.Range("A" & TopNextReport & ":O" & BottomNextReport + TopNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        .ActiveSheet.Range("A" & TopNextReport & ":" & (Chr(Asc("A") + I - 1)) & BottomNextReport + TopNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        ExcApp2.CutCopyMode = False
        ExcDoc2.Close
        Set ExcDoc2 = Nothing
        TopNextReport = BottomNextReport + TopNextReport + 1

        Set ExcApp2 = CreateObject("Excel.Application") 'apre il terzo prospetto
        Set ExcDoc2 = ExcApp2.Workbooks.Open(CashTargetWithChannelByCurrencyReport)
        ExcApp2.Visible = False
        I = 1
        While ExcDoc2.ActiveSheet.Cells(I, 1) = ""
            I = I + 1
        Wend
        BottomNextReport = I - 1

        I = 1
        While InStr(1, ExcDoc2.ActiveSheet.Cells(4, I), "TOTAL Q") = 0
            I = I + 1
        Wend
        I = I - 1
        'ExcDoc2.ActiveSheet.Range("B1:P" & BottomNextReport).Copy
        ExcDoc2.ActiveSheet.Range("B1:" & Chr(Asc("A") + I) & BottomNextReport).Copy

        '.ActiveSheet.Range("G11:U" & 10 + i).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        '    .ActiveSheet.Range("A" & TopNextReport & ":O" & TopNextReport + BottomNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        .ActiveSheet.Range("A" & TopNextReport & ":" & (Chr(Asc("A") + I - 1)) & TopNextReport + BottomNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        ExcApp2.CutCopyMode = False
        ExcDoc2.Close
        Set ExcDoc2 = Nothing

        TopNextReport = TopNextReport + BottomNextReport + 1
        Set ExcApp2 = CreateObject("Excel.Application") 'apre il quarto prospetto
        Set ExcDoc2 = ExcApp2.Workbooks.Open(MainCashTargetReportInUSD)
        ExcApp2.Visible = True
        I = 1
        While ExcDoc2.ActiveSheet.Cells(I, 1) <> ""
            I = I + 1
        Wend
        BottomNextReport = 9

        I = 1
        TopNextReport = TopNextReport + 1
        '    While InStr(1, ExcDoc2.ActiveSheet.Cells(4, i), "TOTAL Q") = 0
        '       i = i + 1
        '  Wend
        ' i = i - 1
        'ExcDoc2.ActiveSheet.Range("B1:P" & BottomNextReport).Copy
        ExcDoc2.ActiveSheet.Range("A1:" & Chr(Asc("E")) & BottomNextReport).Copy

        '.ActiveSheet.Range("G11:U" & 10 + i).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        '    .ActiveSheet.Range("A" & TopNextReport & ":O" & TopNextReport + BottomNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        .ActiveSheet.Range("A" & TopNextReport & ":" & (Chr(Asc("E"))) & TopNextReport + BottomNextReport - 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False
        ExcApp2.CutCopyMode = False
        ExcDoc2.Close
        Set ExcDoc2 = Nothing





        ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit

        With ExcDoc.Worksheets("Sheet1").PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With

    End With

    ExcApp.CutCopyMode = False
    ReportName = GetPathExcelDirectory() & "MergedCashTargetReports.xls"
    ExcApp.Application.DisplayAlerts = False
    DoEvents
    ExcApp.ActiveWorkbook.SaveAs FileName:=ReportName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False


    'ExcApp.Application.DisplayAlerts = False
    ExcDoc.ActiveSheet.Range("A1:E" & TopNextReport + BottomNextReport).Copy
    On Error Resume Next
    'Set oPic = PastePicture(xlBitmap)
    On Error GoTo 0

    If oPic Is Nothing Then
        MsgBox "no image in clipboard"
    Else
        sFilter = "Windows Bitmap (*.bmp),*.bmp"
        vFile = GetPathImages() & "MergedCashTargetReports.BMP"
        If vFile <> False Then
            SavePicture oPic, vFile
        End If
    End If

    ExcDoc.Close
    ExcApp.Application.DisplayAlerts = True

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    Set ExcDoc2 = Nothing
    Set ExcApp2 = Nothing
    MergeAllCashTargetRerports = ReportName
End Function

Function ConvertUTIF8Characters(S As Variant) As Variant
    'If InStr(1, S, "3878258") Then
    '   S = S
    'End If
    If InStr(1, S, "ÑÐ½Ð²") Then
        S = Replace(S, "ÑÐ½Ð²", "???")
    End If
    If InStr(1, S, "Ä") Then
        S = Replace(S, "Ä", "")
    End If
    If InStr(1, S, "ÐŸÐ—") Then
        S = Replace(S, "ÐŸÐ—", "??")
    End If

    If InStr(1, S, "Ð¼Ñ‹ÑˆÐ¸") Then
        S = Replace(S, "Ð¼Ñ‹ÑˆÐ¸", "????")
    End If

    If InStr(1, S, "Ð°Ð²Ð³") Then
        S = Replace(S, "Ð°Ð²Ð³", "???")
    End If
    If InStr(1, S, "Ð¾Ð±Ñ‰Ð¸Ð¹") Then
        S = Replace(S, "Ð¾Ð±Ñ‰Ð¸Ð¹", "?????")
    End If
    If InStr(1, S, "Ð¼Ð°Ð¹") Then
        S = Replace(S, "Ð¼Ð°Ð¹", "???")
    End If

    If InStr(1, S, "Ð´") Then
        S = Replace(S, "Ð´", "?")
    End If
    If InStr(1, S, "Ð°") Then
        S = Replace(S, "Ð°", "?")
    End If
    If InStr(1, S, "°Ð") Then
        S = Replace(S, "°Ð", "?")
    End If
    If InStr(1, S, "·Ð") Then
        S = Replace(S, "·Ð", "?")
    End If
    If InStr(1, S, "Ã³") Then
        S = Replace(S, "Ã³", "r")
    End If
    If InStr(1, S, "Å¼") Then
        S = Replace(S, "Å¼", "o")
    End If
    If InStr(1, S, "Å„") Then
        S = Replace(S, "Å„", "e")
    End If
    If InStr(1, S, "Å‚") Then
        S = Replace(S, "Å‚", "l")
    End If
    If InStr(1, S, "Ãˆ") Then
        S = Replace(S, "Ãˆ", "E")
    End If
    If InStr(1, S, "Ã¼") Then
        S = Replace(S, "Ã¼", "ü")
    End If
    If InStr(1, S, "Ðœ") Then
        S = Replace(S, "Ðœ", "M")
    End If
    If InStr(1, S, "ÄŸ") Then
        S = Replace(S, "ÄŸ", "g")
    End If
    If InStr(1, S, "Ð¤") Then
        S = Replace(S, "Ð¤", "?")
    End If
    If InStr(1, S, "Ð¾") Then
        S = Replace(S, "Ð¾", "?")
    End If
    If InStr(1, S, "Ðº") Then
        S = Replace(S, "Ðº", "?")
    End If
    If InStr(1, S, " Ñ") Then
        S = Replace(S, " Ñ", "?")
    End If
    If InStr(1, S, "ºÐ") Then
        S = Replace(S, "ºÐ", "?")
    End If
    If InStr(1, S, "Ð¾") Then
        S = Replace(S, "Ð¾", "?")
    End If
    If InStr(1, S, "Ð½") Then
        S = Replace(S, "Ð½", "?")
    End If
    If InStr(1, S, "Ð¡") Then
        S = Replace(S, "Ð¡", "?")
    End If
    If InStr(1, S, "Ðž") Then
        S = Replace(S, "Ðž", "?")
    End If
    If InStr(1, S, " Ð") Then
        S = Replace(S, " Ð", "?")
    End If
    If InStr(1, S, " Ñ") Then
        S = Replace(S, " Ñ", "?")
    End If
    If InStr(1, S, "ƒÑ") Then
        S = Replace(S, "ƒÑ", "?")
    End If
    If InStr(1, S, "ÑÐ") Then
        S = Replace(S, "ÑÐ", "?")
    End If
    If InStr(1, S, "Ñ,") Then
        S = Replace(S, "Ñ,", "?")
    End If
    If InStr(1, S, Chr(209) & Chr(129)) Then
        S = Replace(S, Chr(209) & Chr(129), "?")
    End If
    If InStr(1, S, Chr(209) & Chr(130)) Then
        S = Replace(S, Chr(209) & Chr(130), "?")
    End If
    If InStr(1, S, Chr(208) & Chr(158)) Then
        S = Replace(S, Chr(208) & Chr(158), "?")
    End If
    If InStr(1, S, Chr(208) & Chr(149)) Then
        S = Replace(S, Chr(208) & Chr(149), "?")
    End If
    If InStr(1, S, Chr(208) & Chr(156)) Then
        S = Replace(S, Chr(208) & Chr(156), "?")
    End If
    If InStr(1, S, "Ä±") Then
        S = Replace(S, "Ä±", "?")
    End If
    If InStr(1, S, "ÅŸ") Then
        S = Replace(S, "ÅŸ", "?")
    End If
    If InStr(1, S, Chr(208) & Chr(158)) Then
        S = Replace(S, Chr(208) & Chr(158), "?")
    End If
    If InStr(1, S, Chr(208) & Chr(161)) Then
        S = Replace(S, Chr(208) & Chr(161), "?")
    End If
    If InStr(1, S, Chr(239) & Chr(191) & Chr(189)) Then
        S = Replace(S, Chr(239) & Chr(191) & Chr(189), "?")
    End If
    If InStr(1, S, "Ð") Then
        S = Replace(S, "Ð", "?")
    End If
    If InStr(1, S, "¿Ñ€") Then
        S = Replace(S, "¿Ñ€", "?")
    End If
    ConvertUTIF8Characters = S
End Function

Sub SetFontColor(ExcelFile As Variant, Coordinates As String, Colour As Long, Optional aThemeColor, Optional aTintAndShade As Variant)
    With ExcelFile
        If Not IsNull(Colour) Then
            .Range(Coordinates).Font.color = Colour
        Else
            .Range(Coordinates).Font.ThemeColor = aThemeColor
            .Range(Coordinates).Font.TintAndShade = aTintAndShade
        End If
    End With
End Sub

Function MergeTXTFiles(PathArray As Variant, Optional LinesToSkip As Integer)
    Dim I, a As Integer
    Dim S As String
    Dim SourceNum As Integer
    Dim DestNum As Integer
    Dim LineToBeWritten As String
    MergeTXTFiles = ""
    S = "C:\Users\" & fOSUserName() & "\DOCUMENTS\MergedFile.txt"

    Open S For Output As #1
    For I = 0 To UBound(PathArray)
        Open PathArray(I) For Input As #2
        While Not EOF(2)
            If (LinesToSkip > 0) And (I > 0) Then
                For a = 1 To LinesToSkip
                    Line Input #2, LineToBeWritten
                Next a
            Else
                Line Input #2, LineToBeWritten
            End If
            Print #1, LineToBeWritten
        Wend
        Close #2
    Next I
    Close #1
    MergeTXTFiles = S
End Function

Sub FixCashReport(ByRef MyXl As Variant)
    Dim row As Integer
    Dim FiscalMonths() As Integer
    Dim Groups() As String
    Dim item As Integer
    Dim Found As Boolean
    Dim FoundCashReceipts, FoundCashTarget As Boolean
    Dim StartingLine As Integer
    Dim Nmonths, NGroups As Integer
    row = 2
    ReDim FiscalMonths(0)
    ReDim Groups(0)
    With MyXl.Worksheets(1)
        While .Cells(row, 1) <> ""

            ' count number of fiscal months
            Found = False
            For item = 0 To UBound(FiscalMonths)
                If FiscalMonths(item) = .Cells(row, 2) Then
                    Found = True
                End If
            Next item
            If Not Found Then
                ReDim Preserve FiscalMonths(UBound(FiscalMonths) + 1)
                FiscalMonths(UBound(FiscalMonths)) = .Cells(row, 2)
            End If


            ' count number of groups (currency or sales channel)
            Found = False
            For item = 0 To UBound(Groups)
                If Groups(item) = .Cells(row, 7) Then
                    Found = True
                End If
            Next item
            If Not Found Then
                ReDim Preserve Groups(UBound(Groups) + 1)
                Groups(UBound(Groups)) = .Cells(row, 7)
            End If

            row = row + 1
        Wend
        StartingLine = row

        For Nmonths = 1 To UBound(FiscalMonths)
            For NGroups = 1 To UBound(Groups)
                FoundCashReceipts = False
                FoundCashTarget = False
                For row = 2 To StartingLine - 1
                    If (.Cells(row, 2) = FiscalMonths(Nmonths)) And (.Cells(row, 7) = Groups(NGroups)) And (.Cells(row, 4) = "Cash Receipts to date (actual)") Then
                        FoundCashReceipts = True
                    End If
                    If (.Cells(row, 2) = FiscalMonths(Nmonths)) And (.Cells(row, 7) = Groups(NGroups)) And (.Cells(row, 4) = "Cash Target") Then
                        FoundCashTarget = True
                    End If
                Next row
                If Not FoundCashTarget Then
                    .Cells(StartingLine, 1) = .Cells(StartingLine - 1, 1)
                    .Cells(StartingLine, 2) = FiscalMonths(Nmonths)
                    .Cells(StartingLine, 3) = .Cells(StartingLine - 1, 3)
                    .Cells(StartingLine, 4) = "Cash Target"
                    '.Cells(StartingLine, 5) = "01-" & .Cells(StartingLine, 1) & "-" & FiscalMonths(Nmonths)
                    .Cells(StartingLine, 5) = "=date(" & .Cells(StartingLine, 1) & "," & FiscalMonths(Nmonths) & ",01)"
                    .Cells(StartingLine, 6) = 0
                    .Cells(StartingLine, 7) = Groups(NGroups)

                    StartingLine = StartingLine + 1
                End If
                If Not FoundCashReceipts Then
                    .Cells(StartingLine, 1) = .Cells(StartingLine - 1, 1)
                    .Cells(StartingLine, 2) = FiscalMonths(Nmonths)
                    .Cells(StartingLine, 3) = .Cells(StartingLine - 1, 3)
                    .Cells(StartingLine, 4) = "Cash Receipts to date (actual)"
                    '.Cells(StartingLine, 5) = "01-" & .Cells(StartingLine, 1) & "-" & FiscalMonths(Nmonths)
                    .Cells(StartingLine, 5) = "=date(" & .Cells(StartingLine, 1) & "," & FiscalMonths(Nmonths) & ",01)"
                    .Cells(StartingLine, 6) = 0
                    .Cells(StartingLine, 7) = Groups(NGroups)

                    StartingLine = StartingLine + 1
                End If

            Next NGroups
        Next Nmonths
    End With

End Sub

Function FillGeneralCashTargetByEmailInUSD() As String
    Dim qdfNew As DAO.QueryDef
    Dim SQLText As String
    Dim rst2 As Variant
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row, col, I, TotalQuarter As Integer
    Dim CashTargetDate As Date
    'Dim UDSExchangeRate As Currency
    FillGeneralCashTargetByEmailInUSD = ""


    With CurrentDb

        '   UDSExchangeRate = DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='USD'")

        Set rst2 = CurrentDb.OpenRecordset("SELECT TOP 1 Tbl_MonthEnd.LatestPaymentDate, Tbl_MonthEnd.FiscalYear, Tbl_MonthEnd.FiscalQuarter, Tbl_MonthEnd.FiscalMonth FROM Tbl_MonthEnd WHERE (((Tbl_MonthEnd.MonthEnd)>=#" & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "mm/dd/yy") & "#)) ;")

        SQLText = "SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, Round(Sum([Tbl_CashCollected].[AmountInUSD]/1000)) AS TotalAmount FROM Tbl_CashCollected GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
                  "UNION SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round(([Tbl_Cash_Target_Breakdown].[AmountInUSD]/1000)) AS TotalAmount FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round(([Tbl_Cash_Target_Breakdown].[amount]/[Tbl_Cash_Target_Breakdown].[AmountInUSD]/1000)) HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"

        SQLText = "SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, Round(Sum([Tbl_CashCollected].[AmountInUSD]/1000)) AS TotalAmount FROM Tbl_CashCollected GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
                  "UNION SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round(([Tbl_Cash_Target_Breakdown].[AmountInUSD]/1000)) AS TotalAmount FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round(([Tbl_Cash_Target_Breakdown].[AmountInUSD]/1000)) HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"


        '    SqlText = "SELECT Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)' AS Description, DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) AS StartDate, (Round(Sum([Tbl_CashCollected].[amount]/1000))/" & UDSExchangeRate & ") AS TotalAmount FROM Tbl_CashCollected GROUP BY Tbl_CashCollected.FiscalYear, Tbl_CashCollected.FiscalMonth, Tbl_CashCollected.FiscalQuarter, 'Cash Receipts to date (actual)', DateValue('01/' & [Tbl_CashCollected].[fiscalmonth] & '/' & [Tbl_CashCollected].[fiscalyear]) HAVING (((Tbl_CashCollected.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_CashCollected.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & ")); " & _
        '             "UNION SELECT Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target' AS Description, DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]) AS StartDate, Round((([Tbl_Cash_Target_Breakdown].[Amount] / [Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency] / 1000)) / " & UDSExchangeRate & ") As TotalAmount FROM Tbl_Cash_Target_Breakdown GROUP BY Tbl_Cash_Target_Breakdown.FiscalYear, Tbl_Cash_Target_Breakdown.FiscalMonth, Tbl_Cash_Target_Breakdown.FiscalQuarter, 'Cash Target', DateValue('01/' & [fiscalmonth] & '/' & [fiscalyear]), Round ((([Tbl_Cash_Target_Breakdown].[Amount] / [Tbl_Cash_Target_Breakdown].[ExchangeRateToMainCurrency] / 1000)) / " & UDSExchangeRate & ") HAVING (((Tbl_Cash_Target_Breakdown.FiscalYear)=" & rst2.Fields("FiscalYear") & ") AND ((Tbl_Cash_Target_Breakdown.FiscalQuarter)=" & rst2.Fields("FiscalQuarter") & "));"

        Set qdfNew = .CreateQueryDef("Query1", SQLText)
        DoCmd.OutputTo acOutputQuery, "Query1", acFormatXLS, GetPathExcelDirectory() & "GeneralCashTargetInUSD.xls", False
        .QueryDefs.Delete ("Query1")

        Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
        Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory() & "GeneralCashTargetInUSD.xls")
        ExcApp.Visible = False
        Rem ExcApp.visible = True
        With ExcDoc
            row = 1
            While .ActiveSheet.Cells(row + 1, 1) <> ""
                row = row + 1
            Wend
            .Sheets.Add
            ExcApp.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                                     "Query1!R1C1:R" & row & "C6", Version:=xlPivotTableVersion10).CreatePivotTable _
                                                     TableDestination:="Sheet1!R3C1", tableName:="PivotTable1", DefaultVersion _
                                                     :=xlPivotTableVersion10
            .Sheets("Sheet1").Select
            row = 3
            col = 1
            With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("Description")
                .Orientation = xlRowField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable1").AddDataField .ActiveSheet.PivotTables( _
                                                                 "PivotTable1").PivotFields("TotalAmount"), "Sum of TotalAmount", xlSum
            With ExcDoc.ActiveSheet.PivotTables("PivotTable1").PivotFields("StartDate")
                .Orientation = xlColumnField
                .Position = 1
            End With
            .ActiveSheet.PivotTables("PivotTable1").PivotFields("Description").AutoSort _
        xlDescending, "Description"
            .ActiveSheet.Range("B5:E6").NumberFormat = "#,##0"
            .ActiveSheet.PivotTables("PivotTable1").ColumnGrand = False
            .ActiveSheet.Range("A1:IV100").Copy
            .ActiveSheet.Range("A1:IV100").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                        :=False, Transpose:=False

            For I = 1 To 3
                If .ActiveSheet.Cells(row + 1, I + 1) <> "" Then
                    If IsDate(.ActiveSheet.Cells(row + 1, I + 1)) Then
                        ' .ActiveSheet.Cells(Row + 1, i + 1) = DateAdd("yyyy", -1, .ActiveSheet.Cells(Row + 1, i + 1))
                    End If
                    .ActiveSheet.Cells(row + 1, I + 1).NumberFormat = """M" & I & """ mmm yy"
                End If
            Next I

            With .ActiveSheet.Range("A1:IV100")
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End With
            .ActiveSheet.Cells(row, col) = "Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            .ActiveSheet.Cells(row + 1, col) = DLookup("Area", "tblGeneral") & " TOTAL                               " & "(USD)"

            .ActiveSheet.Cells(row, col + 1) = ""
            .ActiveSheet.Range("b" & row + 1 & ":d" & row + 1).Copy
            .ActiveSheet.Range("b" & row & ":d" & row).PasteSpecial
            .ActiveSheet.Cells(row + 1, col + 1) = ""
            .ActiveSheet.Cells(row + 1, col + 2) = ""
            .ActiveSheet.Cells(row + 1, col + 3) = ""
            .ActiveSheet.Cells(row + 1, col + 4) = ""
            I = 1
            While .ActiveSheet.Cells(row, I) <> ""
                I = I + 1
            Wend
            '        i = i - 1
            .ActiveSheet.Cells(row, I) = "TOTAL Q" & rst2.Fields("FiscalQuarter") & " FY" & Right(rst2.Fields("FiscalYear"), 2)
            TotalQuarter = I
            Call BoldLetter(ExcApp, "A1:Z4")
            row = row + 4
            .ActiveSheet.Cells(row, 1) = "Actual Performance to date"
            For I = 1 To 4
                If .ActiveSheet.Cells(row - 1, col + I) <> "" Then
                    If (.ActiveSheet.Range(Chr(65 + I) & row - 2) = "") Or (.ActiveSheet.Range(Chr(65 + I) & row - 2) = 0) Then
                        .ActiveSheet.Cells(row, col + I) = "-"
                        .ActiveSheet.Cells(row, col + I).HorizontalAlignment = xlRight
                    Else
                        .ActiveSheet.Cells(row, col + I) = "=" & Chr(65 + I) & row - 1 & "/" & Chr(65 + I) & row - 2
                    End If
                    .ActiveSheet.Cells(row, col + I).NumberFormat = "0%"
                End If
            Next I
            Call BoldLetter(ExcApp, "A" & row & ":Z" & row)
            .ActiveSheet.Range("A4..Z4").Insert Shift:=xlDown
            .ActiveSheet.Range("A6..Z6").Insert Shift:=xlDown
            Call FillCells(ExcApp, "A5:A5", 0, xlThemeColorAccent5, 0.599993896298105)
            I = 1
            While .ActiveSheet.Cells(3, I + 1) <> ""
                I = I + 1
            Wend
            Call FillCells(ExcApp, "A2:" & Chr(64 + I) & "4", 5296274)
            ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        End With
        Call SetFontColor(ExcApp, "A7:E7", -4165632)
        Call SetFontColor(ExcApp, "A8:E9", -6279056)
        Call BoldLetter(ExcApp, "A4:E40")
        ExcDoc.ActiveSheet.Cells(3, 1).HorizontalAlignment = xlCenter
        Call BorderCells(ExcApp, "A" & 2 & ":" & Chr(64 + I) & row + 2)
        Call BorderCells(ExcApp, "A5:" & Chr(64 + I) & row + 2)

        ExcDoc.Worksheets("Sheet1").Range("A1:Z1000").Columns.AutoFit
        ExcDoc.Worksheets("Sheet1").Cells(1, 1) = "GENERAL CASH TARGET"
        '    CashTargetDate = DateAdd("d", -1, Date)
        '   While Weekday(CashTargetDate) = vbSaturday Or Weekday(CashTargetDate) = vbSunday
        '      CashTargetDate = DateAdd("d", -1, CashTargetDate)
        ' Wend
        CashTargetDate = Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "DD MMM YYyy")
        ExcDoc.Worksheets("Sheet1").Cells(1, 1) = "GENERAL CASH TARGET IN USD"
        ExcDoc.ActiveSheet.Range("A1:" & Chr(64 + TotalQuarter) & "1").HorizontalAlignment = xlCenterAcrossSelection

        With ExcApp
            .Application.DisplayAlerts = False
            .ActiveWorkbook.SaveAs FileName:=GetPathExcelDirectory() & "GeneralCashTargetInUSD.xls", FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
            .Application.DisplayAlerts = True
            .Quit
        End With

        Set ExcDoc = Nothing
        Set ExcApp = Nothing
        FillGeneralCashTargetByEmailInUSD = GetPathExcelDirectory() & "GeneralCashTargetInUSD.xls"
    End With

End Function
