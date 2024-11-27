Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mPB_O2 As clsPROGRESS
Private Sub Form_Activate()
    DoCmd.Restore
End Sub
Private Sub UpdateChargebackFileAfterImport()
    Dim rst As Variant
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim row, Count, I As Integer
    Dim FoundChargeback As Boolean
    Dim PreviousDate As Date
    Dim rngXL As Excel.Range

    Rem Chargebacks Open Now
    Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.*, Tbl_Types.ToFillChargbackFile, Tbl_Invoices.Update_date, Tbl_Customers.Name,Tbl_Customers.Credit_controller FROM (Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID) LEFT JOIN Tbl_Customers ON Tbl_Invoices.Customer_ID = Tbl_Customers.Customer_code WHERE (((Tbl_Types.ToFillChargbackFile)=True) AND ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#));")

    If rst.RecordCount > 0 Then
        rst.MoveLast
        rst.MoveFirst

        PreviousDate = Format(FindPreviousDate, "mm/dd/yy")

        Set ExcApp = CreateObject("Excel.Application")
        Set ExcDoc = ExcApp.Workbooks.Open(GetPathChargbackFile & "\chargebacks.xls")
        ExcApp.Visible = False

        row = NumMaxRows(GetPathChargbackFile & "\chargebacks.xls", "", 7)
        With ExcDoc.Sheets(1)
            For I = 7 To row
                Etichetta2.Caption = "Checking chargeback: " & I & " out of " & row
                mPB_O2.Progressione Int(100 / row * I)
                Me.Etichetta7.Caption = CInt(100 / row * I) & "%"
                DoEvents
                If ExcDoc.Sheets(1).Cells(I, 9) = "" Then
                    rst.FindFirst ("Document_Number ='" & .Cells(I, 1) & "'")
                    If rst.NoMatch Then
                        .Cells(I, 9) = Format(Date, "dd-mmm-yyyy")
                    End If
                End If
            Next I

            Rem add new chargebacks
            rst.MoveFirst
            FoundChargeback = False
            Count = 1
            While Not rst.EOF
                Etichetta2.Caption = "Checking chargeback: " & Count & " out of " & rst.RecordCount
                mPB_O2.Progressione Int(Count * 100 / rst.RecordCount)
                Me.Etichetta7.Caption = CInt(Count * 100 / rst.RecordCount) & "%"
                DoEvents
                I = 7
                While (CDbl(.Cells(I, 1)) < rst.Fields("Document_number")) And (I <= row)
                    I = I + 1
                Wend
                If (CDbl(.Cells(I, 1)) <> CDbl(rst.Fields("Document_number"))) Then
                    Set rngXL = ExcDoc.Sheets(1).Rows(I)
                    rngXL.Insert xlShiftDown

                    .Cells(I, 3) = rst.Fields("Customer_ID")
                    .Cells(I, 7) = rst.Fields("Date")
                    .Cells(I, 1) = CDbl(rst.Fields("Document_Number"))
                    .Cells(I, 6) = rst.Fields("Amount")
                    .Cells(I, 5) = rst.Fields("Currency")
                    .Cells(I, 2) = rst.Fields("Name")
                    .Cells(I, 8) = GetNameCreditControllerFromID(rst.Fields("credit_controller"))
                    Set rngXL = Nothing
                End If
                Count = Count + 1
                rst.MoveNext
            Wend
        End With
        ExcDoc.Save
        ExcDoc.Close
        Set ExcApp = Nothing
        Set ExcDoc = Nothing

    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Form_Timer()
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Dim PreviousAttachmentDate As Date
Dim S, StrInputFileName, StrInputFileName2, StrCurrency, Inputdata, CurrencyName, StrFilter As String
Dim CountryCode, DocumentName As String
Dim SqlText As String
Dim qdfNew As DAO.QueryDef
Dim RstInvoices, RstInvoices2, rst, rst2, rstCredit_controller, RstInvoicesToErase, bb As Recordset
Dim RstMissingCurrencies As Recordset
Dim RstTimezone As Recordset
Dim BC, SDF As Recordset
Dim LastInvoiceID As Long
Dim Cancellati, TotLine, TotRec, OLDCustomer_ID, Customer_ID, I, b, r, a, Totpages, Page, Group, NInvoices, CreditController As Integer
Dim NCustomers As Long
Dim Customer_data(1, 11), Invoicestypes(1, 10), Invoices(100000, 10)
Dim Index As Currency
Dim Lastdate, DateBeforeLast As Date
Dim DateTimezone1, DateTimezone2 As Date
Dim StringDateTimeZone As String
Dim LFile, CountFile As Long
Dim ProgressPercentage As Integer
Dim Startdate, Enddate As Date
Dim DateToBeCancelled As Date
Dim Count As Integer

StrFilter = ahtAddFilterItem(StrFilter, "Text files(*.txt)", "*.txt")
StrInputFileName2 = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select an input file...", _
                Flags:=ahtOFN_HIDEREADONLY Or ahtOFN_ALLOWMULTISELECT Or ahtOFN_EXPLORER)

If IsArray(StrInputFileName2) Then
    StrInputFileName = MergeTXTFiles(StrInputFileName2) 'merge more than one statement if the case
Else
    StrInputFileName = StrInputFileName2
End If

If StrInputFileName <> "" Then
    Startdate = Now()
    Label10.Caption = "Start time: " & Format(Startdate, "dd-mmm-yy hh:mm:ss")
    DBEngine.SetOption dbMaxLocksPerFile, 200000
    Me.Testo8.SetFocus
    Me.Etichetta7.Visible = True
    Set mPB_O2 = New clsPROGRESS
    mPB_O2.CreaPB Me.shPB_O2, Orizzontale
    Me.Repaint
    Call SetImportingProcess(True)
    Open StrInputFileName For Input As #1    ' Apre il file per l'input.
    Totpages = 0
    Group = 0
    NInvoices = 0
    Page = 1
    OLDCustomer_ID = 0
    Cancellati = 0
    ProgressPercentage = 0
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Types", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        I = 1
        While Not .EOF
            Invoicestypes(1, I) = rst.Fields.item("Descripition")
            .MoveNext
            I = I + 1
        Wend
        .Close
    End With

    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    Set rstCredit_controller = New ADODB.Recordset
    With rstCredit_controller
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Countries", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    Set RstInvoices = New ADODB.Recordset
    With RstInvoices
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Invoices", , adOpenKeyset, adLockOptimistic, adCmdTable
        If RstInvoices.RecordCount > 0 Then
            LastInvoiceID = DMax("ID", "Tbl_Invoices")
        Else
            LastInvoiceID = 0
        End If
    End With

    Set RstInvoicesToErase = New ADODB.Recordset
    With RstInvoicesToErase
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Invoices", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    Set RstTimezone = New ADODB.Recordset
    With RstTimezone
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Timezones", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    DoCmd.RunSQL "DELETE Tbl_Invoices.Update_date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)<#" & Format(DateAdd("d", DLookup("MaxOldStatementInArchive", "TblGeneral") * -1, Date), "mm/dd/yy") & "#));"

    CurrencyName = ""
    On Error GoTo CloseFiles
    Index = 0
    CountFile = 0
    LFile = FileLen(StrInputFileName)
    mPB_O2.DefinisciPB

    Rem delete history invoices < 1 year
    DoCmd.RunSQL "Delete Tbl_Invoices_History.Date FROM Tbl_Invoices_History WHERE (((Tbl_Invoices_History.Date)<DateAdd('yyyy',-1,Now())));"

    Rem delete history invoices of all customers
    DoCmd.RunSQL "DELETE Tbl_Invoices_History.Customer_ID, Tbl_Invoices_History.Update_date FROM Tbl_Invoices_History WHERE (((Tbl_Invoices_History.Update_date)=#" & Format(Now(), "mm/dd/yyyy") & "#));"

    While Not EOF(1)
        Index = 0
        While InStr(1, Inputdata, "Page:") = 0
            Line Input #1, Inputdata
            CountFile = CountFile + Len(Inputdata) + 2
        Wend
        Totpages = Trim(Mid$(Inputdata, InStr(1, Inputdata, "Page:") + 11, 3))
            If Page = 1 Then
                With rst
                    Customer_data(1, 1) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 2) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 3) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 4) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 5) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 6) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 7) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 8) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 9) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 10) = Inputdata
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                    Customer_data(1, 11) = Inputdata

                    r = 11
                    While InStr(1, Customer_data(1, r), "Contact:") = 0
                        r = r - 1
                    Wend
                    Customer_ID = CLng(Trim(Mid$(Customer_data(1, r - 2), 100, 20)))
                    CountryCode = Trim(Left$(Customer_data(1, 9), 50))
                    TotRec = rst.RecordCount
                        If TotRec = 0 Then
                            rst.AddNew
                            .Fields.item("Customer_code") = Customer_ID
                            .Fields("NextAppointment") = Now()

                            rstCredit_controller.Find ("Code='" & CountryCode & "'")
                            If rstCredit_controller.EOF Then
                                rstCredit_controller.MoveFirst
                                rstCredit_controller.Find ("Code='" & CountryCode & "'")
                            End If
                            If Not rstCredit_controller.EOF Then
                                .Fields("Credit_controller") = rstCredit_controller.Fields("Credit_controller")
                            End If
                        Else
                            rst.Find ("Customer_code=" & Customer_ID)
                            If rst.EOF Then
                                rst.MoveFirst
                                rst.Find ("Customer_code=" & Customer_ID)
                            End If
                            If rst.EOF Then
                                rst.AddNew
                                .Fields.item("Customer_code") = Customer_ID
                                rstCredit_controller.Find ("Code='" & CountryCode & "'")
                                If rstCredit_controller.EOF Then
                                    rstCredit_controller.MoveFirst
                                    rstCredit_controller.Find ("Code='" & CountryCode & "'")
                                End If
                                If Not rstCredit_controller.EOF Then
                                    .Fields("Credit_controller") = rstCredit_controller.Fields("Credit_controller")
                                End If
                            End If
                        End If
                    .Fields("Country") = CountryCode


                        a = -1
                        For Count = 1 To 11
                            If InStr(1, Customer_data(1, Count), "Attn:") <> 0 Then
                                a = Count
                                Exit For
                            End If
                        Next
                        If a = -1 Then
                            a = 1
                        End If
                    .Fields("Name") = (Trim(Left$(Customer_data(1, a + 1), 50)))
                    .Fields("Address") = (Trim(Left$(Customer_data(1, a + 2), 50)))
                    .Fields("Address2") = (Trim(Left$(Customer_data(1, a + 3), 50)))

                    For b = a + 4 To r - 3
                        .Fields("Address3") = (Trim(Left$(Customer_data(1, b), 50)))
                        .Fields("Address4") = (Trim(Left$(Customer_data(1, b + 1), 50)))
                        .Fields("Address5") = (Trim(Left$(Customer_data(1, b + 2), 50)))
                    Next b

                    .Fields("Update_date") = Now()
                    If Customer_ID = OLDCustomer_ID Then
                        Group = Group + 1
                    Else
                        OLDCustomer_ID = Customer_ID
                        Group = 0
                        .Fields("Index") = 0
                        Set RstInvoicesToErase = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Customer_ID FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Customer_ID & ") And ((Tbl_Invoices.Update_date) = #" & Format(Now(), "mm/dd/yyyy") & "#) AND (Tbl_Invoices.ID<=" & LastInvoiceID & "));")
                        If RstInvoicesToErase.RecordCount > 0 Then
                            Cancellati = Cancellati + RstInvoicesToErase.RecordCount
                            DoCmd.RunSQL "Delete Tbl_Invoices.Customer_ID FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID)=" & Customer_ID & ") AND ((Tbl_Invoices.Update_date)=#" & Format(Now(), "mm/dd/yyyy") & "#) AND (Tbl_Invoices.ID<=" & LastInvoiceID & "));"
                        End If
                    End If

                    If Trim((Mid$(Customer_data(1, 1), 65, 35))) <> "" Then
                        .Fields("OWN_company") = Trim((Mid$(Customer_data(1, 1), 65, 35)))
                        .Fields("OWN_bank_details1") = (Mid$(Customer_data(1, 2), 65, 35))
                        .Fields("OWN_bank_details2") = (Mid$(Customer_data(1, 3), 65, 35))
                        .Fields("OWN_bank_details3") = (Mid$(Customer_data(1, 4), 65, 35))
                        .Fields("OWN_bank_details4") = (Mid$(Customer_data(1, 5), 65, 35))
                    End If
                    .Update
                End With
                While InStr(1, Inputdata, "                                                           Amount") = 0
                    Line Input #1, Inputdata
                    CountFile = CountFile + Len(Inputdata) + 2
                Wend
                Line Input #1, Inputdata
                CountFile = CountFile + Len(Inputdata) + 2
                Line Input #1, Inputdata
                CountFile = CountFile + Len(Inputdata) + 2
                TotLine = 0
                While (InStr(1, Inputdata, "Balance:") = 0) And (InStr(1, Inputdata, "LFINAGEESDC") = 0)
                    Inputdata = ConvertUTIF8Characters(Inputdata & "")
                    If (Len(Inputdata) <> 0) And (Trim(Mid$(Inputdata, 23, 9)) <> "") Then
                        If (TotLine <> 0) And (TotLine Mod 30 = 0) Then
                            While (InStr(1, Inputdata, "                                                           Amount") = 0)
                                Line Input #1, Inputdata
                                CountFile = CountFile + Len(Inputdata) + 2
                            Wend
                            Line Input #1, Inputdata
                            CountFile = CountFile + Len(Inputdata) + 2
                            Line Input #1, Inputdata
                            CountFile = CountFile + Len(Inputdata) + 2
                        End If
                        If Trim(Mid$(Inputdata, 23, 9)) <> "" Then
                            TotLine = TotLine + 1
                            Invoices(TotLine, 1) = Trim(Left$(Inputdata, 11)) 'invoicenumber

                            Invoices(TotLine, 2) = Mid$(Inputdata, 12, 9) 'date
                            TotRec = 1
                            DocumentName = Trim(Mid$(Inputdata, 21, 14))
                            While DocumentName <> Invoicestypes(1, TotRec)
                                TotRec = TotRec + 1
                            Wend

                            Invoices(TotLine, 3) = TotRec 'type
                            Invoices(TotLine, 4) = Trim(Mid$(Inputdata, 34, 26)) 'cust reference
                            Invoices(TotLine, 5) = CCur(Trim((Mid$(Inputdata, 74, 17)))) ' value
                            Invoices(TotLine, 6) = Trim(Mid$(Inputdata, 88, 13)) 'due date
                            Invoices(TotLine, 7) = CCur(Trim(Mid(Inputdata, 61, 13))) ' Original transaction amount
                            Invoices(TotLine, 8) = Trim(Mid(Inputdata, 110, 11)) 'SO number
                            Invoices(TotLine, 9) = Trim(Mid(Inputdata, 122, 11)) 'Pull ticket number
                            Invoices(TotLine, 10) = Trim(Mid(Inputdata, 101, 8)) 'Factura number

                        Else
                            Invoices(TotLine, 4) = Invoices(TotLine, 4) & " " & Trim(Inputdata)
                        End If
                        Line Input #1, Inputdata
                        CountFile = CountFile + Len(Inputdata) + 2
                    Else
                        'If ((Trim(Mid$(Inputdata, 23, 9)) <> "")) Then
                        If (Len(Inputdata) > 10) And (Trim(Mid$(Inputdata, 12, 9)) = "") Then
                            Invoices(TotLine, 4) = Left(Invoices(TotLine, 4) & Trim(Mid$(Inputdata, 34, 26)), 50)
                            Invoices(TotLine, 8) = Left(Invoices(TotLine, 8) & Trim(Mid$(Inputdata, 111, 10)), 15)
                            Invoices(TotLine, 9) = Left(Invoices(TotLine, 9) & Trim(Mid$(Inputdata, 122, 11)), 18)
                            Invoices(TotLine, 10) = Left(Invoices(TotLine, 10) & Trim(Mid$(Inputdata, 101, 8)), 25)
                            'TotLine = TotLine + 1
                        End If
                        Line Input #1, Inputdata
                        CountFile = CountFile + Len(Inputdata) + 2
                    End If

                Wend
                Rem se inputdata finisce con balance allora ricavare il nome della valuta
                If (InStr(1, Inputdata, "Balance:") <> 0) Then
                    CurrencyName = Trim(Mid$(Inputdata, (InStr(1, Inputdata, "Balance:")) + 9, 3))
                End If
                With RstInvoices
                    For b = 1 To TotLine
                        .AddNew
                        .Fields("Customer_ID") = rst.Fields.item("Customer_code")
                        .Fields("Update_date") = Date
                        .Fields("Date") = Invoices(b, 2)
                        .Fields("Document_Number") = Invoices(b, 1)
                        .Fields("Type") = Invoices(b, 3)
                        .Fields("Amount") = Invoices(b, 5)
                        .Fields("Customer_reference") = Left$(Invoices(b, 4), 50)
                        .Fields("Overdue_Date") = Invoices(b, 6)
                        .Fields("OriginalAmount") = Invoices(b, 7)
                        .Fields("SONumber") = Invoices(b, 8)
                        If Invoices(b, 9) <> "" Then
                            .Fields("PullTicketN#") = Invoices(b, 9)
                        End If
                        If Invoices(b, 10) <> "" Then
                            .Fields("CustomsInvoiceNumber") = Invoices(b, 10)
                        End If
                        If .Fields("Overdue_Date") > Date Then
                            rst.Fields("Index") = rst.Fields("Index") + 1
                        Else
                            If .Fields("Amount") > 0 Then
                                rst.Fields("Index") = rst.Fields("Index") + .Fields("Amount") * (DateDiff("m", .Fields("Overdue_Date"), Date) + 1)
                            Else
                                rst.Fields("Index") = rst.Fields("Index") + .Fields("Amount")
                            End If
                        End If
                        rst.Update
                        .Update
                        NInvoices = NInvoices + 1
                        Etichetta2.Caption = "Invoices imported: " & NInvoices
                        Etichetta3.Caption = "Invoices cancelled: " & Cancellati
                        If ProgressPercentage <> (CInt(100 * CountFile / LFile)) Then
                            ProgressPercentage = (CInt(100 * CountFile / LFile))
                            mPB_O2.Progressione (ProgressPercentage)
                            Me.Etichetta7.Caption = ProgressPercentage & "%"
                            Me.Repaint
                            DoEvents
                        End If
                    Next b
                End With
                Rem mettere qui una query che inserisce il nome della valuta

                If CurrencyName <> "" Then
                    CurrentDb.Execute ("UPDATE Tbl_Invoices SET Tbl_Invoices.[Currency] = '" & CurrencyName & "' WHERE (((Tbl_Invoices.Customer_ID)=" & Customer_ID & ") AND ((Tbl_Invoices.[Currency]) Is Null) AND ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yyyy") & "#));")
                    CurrencyName = ""
                End If

            End If
            While InStr(1, Inputdata, "Page:") = 0
                Line Input #1, Inputdata
                CountFile = CountFile + Len(Inputdata) + 2
            Wend
        Rem fine cliente
         rst.Update
    Wend
    GoTo CloseFiles
CloseFiles:
    Close #1
    mPB_O2.Progressione (0)
    rst.Update
    rst.Close
    rstCredit_controller.Close
    RstInvoices.Close
    If NInvoices = 0 Then
        b = MsgBox("The number of invoices imported is 0; weird isn't it ? Please check the statement later. The import has been stopped now.", vbOKOnly)
        Exit Sub
    End If
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
        .Fields("Lastupdate") = Now()
        .Update
        .Close
    End With


    Rem ################ FIND DATE BEFORE LAST UPDATE ################
    Set rst = New Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Invoices", , adOpenKeyset, adLockOptimistic, adCmdTable
        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date FROM Tbl_Invoices GROUP BY Tbl_Invoices.Update_date;")
        rst.MoveLast
        Lastdate = .Fields("Update_date")
        .MovePrevious
        DateBeforeLast = FindPreviousDate
        .Close
    End With


    CreditController = GetNumCreditController(fOSUserName)

    Set RstInvoices = New ADODB.Recordset
    With RstInvoices
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Invoices", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    Set RstInvoices2 = New ADODB.Recordset
    With RstInvoices2
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Invoices", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    I = 0

    Rem ######### MARK CUSTOMER TO SEND NEW STATEMENT (WHO HAS INVOICES MISSING) ##############
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable
        Set rst = CurrentDb.OpenRecordset("SELECT * FROM Tbl_Customers;")
        .MoveLast
        .MoveFirst
    End With

    Rem UPDATE QUERIES IN NEW INVOICES

    Etichetta2.Caption = ""
    Etichetta3.Caption = ""
    Set RstInvoices = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Customer_ID, Tbl_Invoices.Date,Tbl_Invoices.Document_Number, Tbl_Invoices.Query,Tbl_Invoices.memo FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(DateBeforeLast, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Query) >0));")
    Set RstInvoices2 = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Customer_ID, Tbl_Invoices.Date,Tbl_Invoices.Document_Number, Tbl_Invoices.Query,Tbl_Invoices.memo FROM Tbl_Invoices WHERE ((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) ORDER BY Tbl_Invoices.Customer_ID, Tbl_Invoices.Document_Number;")
    mPB_O2.Progressione (0)
    Me.Repaint
    DoEvents
    If RstInvoices.RecordCount > 0 Then
        I = 0
        RstInvoices.MoveLast
        RstInvoices.MoveFirst
        While RstInvoices.EOF = False
            RstInvoices2.FindFirst "Document_Number='" & RstInvoices.Fields("Document_Number") & "' AND customer_id=" & RstInvoices.Fields("customer_id")
            If RstInvoices2.NoMatch = False Then
                I = I + 1
                Etichetta2.Caption = "Updating query " & I & " out of " & RstInvoices.RecordCount
                DoEvents
                RstInvoices2.Edit
                RstInvoices2.Fields("query") = RstInvoices.Fields("query")
                RstInvoices2.Fields("memo") = RstInvoices.Fields("memo")
                RstInvoices2.Update
            Else
                rst.MoveFirst
                rst.FindFirst "Customer_code=" & RstInvoices.Customer_ID & ""
            End If
            Etichetta2.Caption = "Updating query " & I & " out of " & RstInvoices.RecordCount
            mPB_O2.Progressione Int(I * 100 / RstInvoices.RecordCount)
            Me.Etichetta7.Caption = CInt(I * 100 / RstInvoices.RecordCount) & "%"
            Me.Repaint
            DoEvents
            RstInvoices.MoveNext
        Wend
    End If

    Etichetta2.Caption = "Updating chargeback file... wait"
    DoEvents
    Call UpdateChargebackFileAfterImport

    Rem ########## CLOSE TABLES #########
    rst.Close
    RstInvoices.Close
    RstInvoices2.Close
    RstTimezone.Close

    Rem copy new invoices in archive
    CurrentDb.QueryDefs("QueryInvoicesLastDate").SQL = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(FindLastDate, "mm/dd/yy") & "#));"
    CurrentDb.QueryDefs("QueryInvoicesPreviousDate").SQL = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(FindPreviousDate, "mm/dd/yy") & "#));"

    Etichetta2.Caption = "Updating queries new invoices... wait"
    DoEvents
    DoCmd.OpenQuery "QueryNewInvoices"

    Rem delete Historical invocies oldest than one year
    Etichetta2.Caption = "Deleting old historical invoices... wait"
    DoEvents
    DoCmd.OpenQuery "QueryDeleteOldHistoricalInvoices"

    Rem delete all intercompany invocies
    Etichetta2.Caption = "Deleting Intercompany invoices... wait"
    DoEvents
    DoCmd.OpenQuery "QueryDeleteInvoicesIntercompanies"

    Rem delete all Historical intercompany invocies
    Etichetta2.Caption = "Deleting Historical Intercompany Invoices... wait"
    DoEvents
    DoCmd.OpenQuery "QueryDeleteHistoricalInvoicesIntercompanies"

    Rem Update payment date in Invoices history
    Etichetta2.Caption = "Updating payment dates... wait"
    DoEvents
    DoCmd.OpenQuery "QueryUpdatePaymentDateInvoicesHistory"


    Rem ### move place chargeback in Query invoice field ###
    CurrentDb.Execute ("UPDATE Tbl_Invoices SET Tbl_Invoices.Query = 25 WHERE (((Tbl_Invoices.Update_date)=Date()) AND ((Tbl_Invoices.Type)=4) AND ((Tbl_Invoices.Query)=0));")
    Etichetta2.Caption = "Moving chargebacks... wait"
    DoEvents

    Rem DSO calculation
    Etichetta2.Caption = "Calculating DSO... wait"
    DoEvents

    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With

    Set RstInvoices = CurrentDb.OpenRecordset("SELECT Tbl_Customers.*, QueryTotalExposureConvertedInEUR.TotalExposureConvertedInEUR, QueryTotalInvoicedConvertedInEUR.TotalInvoicedConvertedInEUR FROM (Tbl_Customers INNER JOIN QueryTotalExposureConvertedInEUR ON Tbl_Customers.Customer_code = QueryTotalExposureConvertedInEUR.Customer_ID) INNER JOIN QueryTotalInvoicedConvertedInEUR ON QueryTotalExposureConvertedInEUR.Customer_ID = QueryTotalInvoicedConvertedInEUR.Customer_ID;")
    With RstInvoices
        .MoveLast
        .MoveFirst
        I = 1
        mPB_O2.Progressione (0)
        Me.Repaint
        DoEvents
        While .EOF = False
            rst.MoveFirst
            rst.Find ("Customer_code=" & RstInvoices.Fields("Customer_code"))
            If .Fields("TotalInvoicedConvertedInEUR") <> 0 Then
                rst.Fields("DSO") = .Fields("TotalExposureConvertedInEUR") / (.Fields("TotalInvoicedConvertedInEUR") / 365)
            Else
                rst.Fields("DSO") = 0
            End If
            Etichetta2.Caption = "Calculating DSO per customer " & I & " out of " & RstInvoices.RecordCount
            mPB_O2.Progressione Int(I * 100 / RstInvoices.RecordCount)
            Me.Etichetta7.Caption = CInt(I * 100 / RstInvoices.RecordCount) & "%"
            Me.Repaint
            DoEvents
            I = I + 1
            .MoveNext
        Wend
    End With
    rst.Update
    rst.Close

    Rem copy one statement per month in Tbl_Historical_Statements
    Call CopyStatement
    Etichetta2.Caption = "Copying Statements for statistic purposes... wait"
    DoEvents

    Rem #########  Updates Historical Credit Limits #######
    Call UpdateHistoricalCL
    Etichetta2.Caption = "Updating Historical Credit Limits... wait"
    DoEvents

    Rem ### move 1 day ahead all customers with no statement ###
  '  CurrentDb.Execute ("UPDATE Tbl_Customers SET Tbl_Customers.NextAppointment = DateAdd('d',+1,Date()) & ' 07:00:00' WHERE (((Exists (SELECT Tbl_Invoices.Customer_ID FROM Tbl_Invoices WHERE Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID And Tbl_Invoices.Update_date =#" & Format(Date, "mm/dd/yy") & "#))=False)) ")

    Rem ### based on the ones from the day before remove the flagtbl_invoices.QueryToBePrinted ###
    Etichetta2.Caption = "Updating queries to be printed in the statement... wait"
    DoEvents
    With CurrentDb
        SqlText = "SELECT Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date, Tbl_Invoices.Date, Tbl_Invoices.Document_Number, Tbl_Invoices.QueryToBePrinted FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=#" & Format(DateAdd("d", -1, Date), "mm/dd/yy") & "#) AND ((Tbl_Invoices.QueryToBePrinted)=No));"
        On Error Resume Next
        .QueryDefs.Delete ("Queryyesterday")
        Set qdfNew = .CreateQueryDef("Queryyesterday", SqlText)

        SqlText = "SELECT Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date, Tbl_Invoices.Date, Tbl_Invoices.Document_Number, Tbl_Invoices.QueryToBePrinted FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#));"
        .QueryDefs.Delete ("QueryToday")
        Set qdfNew = .CreateQueryDef("QueryToday", SqlText)
        On Error GoTo 0

        CurrentDb.Execute ("UPDATE QueryToday LEFT JOIN Queryyesterday ON (QueryToday.Customer_ID = Queryyesterday.Customer_ID) AND (QueryToday.Date = Queryyesterday.Date) AND (QueryToday.Document_Number = Queryyesterday.Document_Number) SET QueryToday.QueryToBePrinted = False WHERE (((Queryyesterday.QueryToBePrinted)=False));")
    End With

    Rem ### based on the ones from the day before add Tbl_Invoices.attachment to the invoices open at the moment in the statement###
    Etichetta2.Caption = "Updating attachments... wait"
    DoEvents
    With CurrentDb
        On Error Resume Next

        If (Format(DMax("Update_date", "tbl_invoices", "Update_date<#" & Format(Date, "mm/dd/yy") & "#"), "mm/dd/yy")) = "" Then
            PreviousAttachmentDate = Format(DateAdd("d", -1, Now()), "mm/dd/yy")
        Else
            PreviousAttachmentDate = Format(DMax("Update_date", "tbl_invoices", "Update_date<#" & Format(Date, "mm/dd/yy") & "#"), "mm/dd/yy")
        End If

        SqlText = "SELECT Tbl_Invoices.Customer_ID, Tbl_Invoices.Document_Number, Tbl_Invoices.Customer_reference, Tbl_Invoices.Type, Tbl_Invoices.Date, Tbl_Invoices.Attachment FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#));"
        Set qdfNew = .CreateQueryDef("Query6", SqlText)
        SqlText = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Attachment, Tbl_Invoices.Customer_ID, Tbl_Invoices.Document_Number, Tbl_Invoices.Customer_reference, Tbl_Invoices.Type, Tbl_Invoices.Date FROM Tbl_Invoices WHERE Tbl_Invoices.Update_date=#" & PreviousAttachmentDate & "# AND Tbl_Invoices.Attachment=True;"
        Set qdfNew = .CreateQueryDef("Query55", SqlText)
        CurrentDb.Execute ("UPDATE Query55 INNER JOIN Query6 ON (Query55.Date = Query6.Date) AND (Query55.Type = Query6.Type) AND (Query55.Customer_reference = Query6.Customer_reference) AND (Query55.Document_Number = Query6.Document_Number) AND (Query55.Customer_ID = Query6.Customer_ID) SET Query6.Attachment = True;")
        .QueryDefs.Delete ("Query55")
        .QueryDefs.Delete ("Query6")

        ' remove invoice attachmnets of closed invoices + clear DB
        SqlText = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Attachment, Tbl_Invoices.Customer_ID, Tbl_Invoices.Document_Number, Tbl_Invoices.Customer_reference, Tbl_Invoices.Type, Tbl_Invoices.Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_Invoices.Attachment)=True));"
        Set qdfNew = .CreateQueryDef("Query6", SqlText)

        SqlText = "SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Attachment, Tbl_Invoices.Customer_ID, Tbl_Invoices.Document_Number, Tbl_Invoices.Customer_reference, Tbl_Invoices.Type, Tbl_Invoices.Date, Left([document_number] & '       ',7) & Left([Date] & '          ',10) & [Type] & Left([Customer_reference] & '               ',15) AS AttachID FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=#" & PreviousAttachmentDate & "#) AND ((Tbl_Invoices.Attachment)=True));"
        Set qdfNew = .CreateQueryDef("Query55", SqlText)
        Set RstInvoices = CurrentDb.OpenRecordset("SELECT Query6.Update_date, Query55.Attachment, Query55.Customer_ID, Query55.Document_Number, Query55.Customer_reference, Query55.Type, Tbl_InvoiceAttachments.DocumentID, Query55.AttachID, Tbl_InvoiceAttachments.AttachName FROM (Query55 LEFT JOIN Query6 ON (Query55.Date = Query6.Date) AND (Query55.Type = Query6.Type) AND (Query55.Customer_reference = Query6.Customer_reference) AND (Query55.Document_Number = Query6.Document_Number) AND (Query55.Customer_ID = Query6.Customer_ID)) LEFT JOIN Tbl_InvoiceAttachments ON Query55.AttachID = Tbl_InvoiceAttachments.DocumentID WHERE (((Query6.Update_date) Is Null));")
        While Not RstInvoices.EOF
            If Not (IsNull(RstInvoices.Fields("AttachName"))) Then
'                Kill (RstInvoices.Fields("AttachName")) 'cancella dal disco il file di attachment
            End If
            CurrentDb.Execute ("DELETE Tbl_InvoiceAttachments.*, Tbl_InvoiceAttachments.DocumentID FROM Tbl_InvoiceAttachments WHERE (((Tbl_InvoiceAttachments.DocumentID)='" & RstInvoices.Fields("AttachID") & "'));") 'cancella il Record in invoiceattachment
            RstInvoices.MoveNext
        Wend
        .QueryDefs.Delete ("Query55")
        .QueryDefs.Delete ("Query6")
        On Error GoTo 0
    End With

    Rem ### se manca valuta(e) in tabella cambi allora manda email di notifica a tutti gli utenti
    Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Currency, Tbl_Currencies.CurrencyID FROM Tbl_Invoices LEFT JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yy") & "#)) GROUP BY Tbl_Invoices.Currency, Tbl_Currencies.CurrencyID HAVING (((Tbl_Currencies.CurrencyID) Is Null));")
    If rst.RecordCount > 0 Then
        S = ""
        rst.MoveFirst
        While Not rst.EOF
            S = S & "," & rst.Fields("currency")
            rst.MoveNext
        Wend
        S = Mid(S, 2, Len(S))
        Call SendEmails(GetAllUserEmails, "", "Access DB - Missing Currencies as of " & Format(Date, "dd-mmm-yyyy"), "The following currencies are missing from the currency archive:" & _
        Chr(13) & S & Chr(13) & Chr(13) & "Please insert them with the relevant exachange rate immediately.", "" _
        , "")
        If rst.RecordCount < 6 Then
            b = MsgBox("The following currencies are missing from the currency archive:" & Chr(13) & S & Chr(13) & Chr(13) & "Please insert them with the relevant exachange rate immediately.", vbExclamation, "Missing Currency")
        Else
            b = MsgBox("A few currencies are missing from the currency archive. An email with all details has been sent. " & Chr(13) & S & Chr(13) & "Please insert them with the relevant exachange rate immediately.", vbExclamation, "Missing Currency")
        End If
    End If
    Set rst = Nothing
    Rem ### se manca ad uno o piu' clienti non e' associato nessun credit controller allora manda email di notifica a tutti gli utenti
    Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Customers.Credit_controller, Tbl_Customers.Name, Tbl_Customers.Customer_code FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller) Is Null Or (Tbl_Customers.Credit_controller)=0));")
    If rst.RecordCount > 0 Then
        S = ""
        rst.MoveFirst
        While Not rst.EOF
            S = S & Chr(13) & rst.Fields("Name") & " Customer ID(" & rst.Fields("Customer_code") & ")"
            rst.MoveNext
        Wend
        Call SendEmails(GetAllUserEmails, "", "Access DB - Customers not associated as of " & Format(Date, "dd-mmm-yyyy"), "The following customers are not associated with any Credit Controller in the Master Data archive:" & _
        Chr(13) & S & Chr(13) & Chr(13) & "Please link them with the relevant Credit Controller immediately.", "" _
        , "")
        If rst.RecordCount < 6 Then
            b = MsgBox("The following customers are not associated with any Credit Controller in the Master Data archive:" & Chr(13) & S & Chr(13) & Chr(13) & "Please link them with the relevant Credit Controller immediately.", vbExclamation, "Customers not associated")
        Else
            b = MsgBox("A few customers are not associated with any Credit Controller in the Master Data archive." & Chr(13) & S & Chr(13) & "Please link them with the relevant Credit Controller immediately.", vbExclamation, "Customers not associated")
        End If
    End If

    Rem *** delete old daily exposure data and add the most updated ones ***
    CurrentDb.Execute ("DELETE Tbl_DailyExposure.*, Tbl_DailyExposure.RefDate FROM Tbl_DailyExposure WHERE (((Tbl_DailyExposure.RefDate)=# " & Format(Now(), "mm/dd/yy") & "#));")
    DateToBeCancelled = DateAdd("d", Date, -365)
    CurrentDb.Execute ("DELETE Tbl_DailyExposure.*, Tbl_DailyExposure.RefDate FROM Tbl_DailyExposure WHERE (((Tbl_DailyExposure.RefDate)<# " & Format(DateToBeCancelled, "mm/dd/yy") & "#));")
    CurrentDb.Execute ("INSERT INTO Tbl_DailyExposure ( CustomerID, RefDate, ARExposure, TotalOverdue, OracleCL, [Currency] ) SELECT QueryDailyExposure.Customer_code, Date() AS Expr1, QueryDailyExposure.ARExposure, QueryDailyExposure.AROverdue, QueryDailyExposure.[CreditLimit in EUR], QueryDailyExposure.Currency FROM QueryDailyExposure;")

    Rem *** update next appointment of customer with nil balance to tomorrow 08:00:01
    CurrentDb.Execute ("UPDATE Tbl_Customers LEFT JOIN Tbl_Invoices ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID SET Tbl_Customers.NextAppointment = #" & Format(DateAdd("d", Date, 1), "mm/dd/yy") & " 08:00:01" & "# WHERE (((Tbl_Invoices.Amount) Is Null));")

    Rem *** update the highest exposure master data field with the highest a/r exposure in the last 365 days ***
    CurrentDb.Execute ("UPDATE Tbl_Customers SET Tbl_Customers.HighestExposure = 0;")
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable
    End With
    Set RstInvoices = CurrentDb.OpenRecordset("SELECT Tbl_DailyExposure.CustomerID, Max(Tbl_DailyExposure.ARExposure) AS MaxOfARExposure FROM Tbl_DailyExposure GROUP BY Tbl_DailyExposure.CustomerID;")
    With RstInvoices
        .MoveLast
        .MoveFirst
        I = 1
        mPB_O2.Progressione (0)
        Me.Repaint
        DoEvents
        While .EOF = False
            rst.MoveFirst
            rst.Find ("Customer_code=" & RstInvoices.Fields("CustomerID"))
            rst.Fields("HighestExposure") = RstInvoices.Fields("MaxOfARExposure")
            Etichetta2.Caption = "Calculating Max exposure per customer " & I & " out of " & RstInvoices.RecordCount
            mPB_O2.Progressione Int(I * 100 / RstInvoices.RecordCount)
            Me.Etichetta7.Caption = CInt(I * 100 / RstInvoices.RecordCount) & "%"
            Me.Repaint
            DoEvents
            I = I + 1
            .MoveNext
        Wend
    End With
    rst.Update
    rst.Close

    Set rst = Nothing
    Call SetImportingProcess(False)
    Enddate = Now()
    Label11.Caption = "End Time: " & Format(Enddate, "dd-mmm-yy   hh:mm:ss")
    Label12.Caption = "Elapsed time: " & Format(Format(Enddate - Startdate), "hh:mm:ss")
    DoEvents

    b = MsgBox("Import completed.", vbOKOnly)
    Set mPB_O2 = Nothing

    Rem ######### SHOW WHO SEND NEW STATEMENTS TO #######
''    a = ExportTextDelimited("QueryStatementCSV", "|")
  ''  a = ExportTextDelimited("QueryCustomersCSV", "|")

End If
Me.TimerInterval = 0
Rem DoCmd.OpenForm ("FormSendStatements")
DoCmd.Close acForm, "MskImportStatement", acSaveNo

End Sub
