Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private mPB_O2 As clsPROGRESS
Dim LFile As Long

Private Sub Form_Timer()
Dim rst, rst2 As Variant
'Dim rst2 As Recordset
Dim rst3 As Variant
Dim rst4 As Variant
Dim CustomerID As Long
Dim PaymentDate, FirstDateFiscalMonth, PreviousMonthEnd As Date
Dim CurrencyID, RETNumber As String
Dim Count, C2, CountAddedRecords As Long
Dim Currencies()   As Variant
Dim provv, Inputdata, StrInputFileName, StrInputFileName2, StrFilter As String
Dim CashTargetWithChannelByEmail As String
Dim CashTargetWithChannelByCurrency As String
Dim CashAmount As Currency
Dim Startdate, Enddate As Date
Dim Schedule, DayOfTheWeek As Byte
Dim TargetReportToBeSent As Boolean
Dim ProgressiveDay As Date
Dim ConsolidatedCashTargetReport  As String
Dim EmailDistributionList2 As String
Dim USDExchangeRate As Currency
Dim AmountInUSD As Currency
Dim GeneralCashTargetInUSD As String
Rem a = ExportTextDelimited("QuerycashCSV", "|")
    StrFilter = ahtAddFilterItem(StrFilter, "Text files(*.txt)", "*.txt")
    StrInputFileName2 = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
        DialogTitle:="Please select the Cash received file to upload...", _
        Flags:=ahtOFN_HIDEREADONLY Or ahtOFN_ALLOWMULTISELECT Or ahtOFN_EXPLORER)

    Startdate = Now()
    Label10.Caption = "Start time: " & Format(Startdate, "dd-mmm-yy   hh:mm:ss")


    If IsArray(StrInputFileName2) Then
        StrInputFileName = MergeTXTFiles(StrInputFileName2) 'merge more than one cash collected report if the case
    Else
        StrInputFileName = StrInputFileName2
    End If

    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CurrentProject.Connection
    rst.Open "Tbl_CashCollected", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Currencies.* FROM Tbl_Currencies;")
    rst2.MoveLast
    rst2.MoveFirst
    ReDim Currencies(rst2.RecordCount, 2)
    Count = 0
    While Not rst2.EOF
        Currencies(Count, 1) = rst2.Fields(0)
        Currencies(Count, 2) = rst2.Fields(1)
        Count = Count + 1
        rst2.MoveNext
    Wend
    Me.TimerInterval = 0

    USDExchangeRate = DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='USD'")

    If StrInputFileName <> "" Then
        Set mPB_O2 = New clsPROGRESS
        mPB_O2.CreaPB Me.shPB_O2, Orizzontale
        mPB_O2.DefinisciPB
        mPB_O2.Progressione (0)

        Me.Etichetta99.Visible = True
        Me.Repaint
        CountFile = 0
        LFile = FileLen(StrInputFileName)
        DoCmd.RunSQL "DELETE Tbl_CashCollected.* FROM Tbl_CashCollected WHERE (((Tbl_CashCollected.[Payment Date])<#" & Format(DateAdd("d", -365, Date), "mm/dd/yy") & "#));"
        DoCmd.RunSQL "UPDATE Tbl_CashCollected SET Tbl_CashCollected.PaymentStillAvailable = False;"
        ' reset tick. It will be useful later if some payments disappears from the last upload
        '(for instance when payments are cancelled and replaced with new RET numbers.
        'In this case an email is sent out to all Access user and after investigation one will decide if
        'to cancel the payment in Access too



        Open StrInputFileName For Input As #1    ' Apre il file per l'input.

        Line Input #1, Inputdata
        Count = Len(Inputdata) + 2

        Set rst3 = CurrentDb.OpenRecordset("SELECT Tbl_CashCollected.CustomerID, Tbl_CashCollected.RETnumber FROM Tbl_CashCollected;")

        CountAddedRecords = 0

        PreviousMonthEnd = DateAdd("d", -1, Date)
        While (Weekday(PreviousMonthEnd) = 1) Or (Weekday(PreviousMonthEnd) = 7)
            PreviousMonthEnd = DateAdd("d", -1, PreviousMonthEnd)
        Wend
        PreviousMonthEnd = DMax("MonthEnd", "Tbl_MonthEnd", "MonthEnd<=#" & DMax("MonthEnd", "Tbl_MonthEnd", "MonthEnd<#" & Format(PreviousMonthEnd, "mm/dd/yy") & "#") & "#")

        While Not EOF(1)
            If IsDate(Left(Inputdata, 12)) Then
                PaymentDate = CDate(Left(Inputdata, 12)) ' current memorize payment date
                provv = Inputdata
                Inputdata = ConvertUTIF8Characters(provv & "")
                RETNumber = Trim(Mid(Inputdata, 56, 12))
            ElseIf (InStr(1, Inputdata, "Applied") = 1) Or (InStr(1, Inputdata, "Unapplied") = 1) Then
                    CustomerID = CLng(Trim(Mid(Inputdata, 46, 8)))
                    CashAmount = Trim(Mid(Inputdata, 33, 12))
                    Set rst3 = CurrentDb.OpenRecordset("SELECT Tbl_CashCollected.CustomerID, Tbl_CashCollected.RETnumber, Tbl_CashCollected.[Original amount], Tbl_CashCollected.PaymentStillAvailable, Tbl_CashCollected.[Payment Date] FROM Tbl_CashCollected WHERE CustomerID =" & CustomerID & " AND RETnumber='" & RETNumber & "'" & " AND [Original amount] =" & CashAmount)
                    If rst3.RecordCount = 0 Then
                        With rst
                            .AddNew
                            .Fields("CustomerID") = CustomerID
                            .Fields("Payment Date") = PaymentDate
                            .Fields("currency") = CurrencyID
                            .Fields("ORIGINAL Amount") = CCur(Trim(Mid(Inputdata, 33, 13)))
                            .Fields("Amount") = .Fields("ORIGINAL Amount") * Currencies(C2, 2) 'amount in EUR
                            .Fields("AmountInUSD") = .Fields("Amount") / USDExchangeRate
                            '-------------------------------
                            .Fields("RETnumber") = RETNumber
                            .Fields("PaymentStillAvailable") = True
                            .Update
                            CountAddedRecords = CountAddedRecords + 1
                            Label1.Caption = "Cash receipts added: " & CountAddedRecords
                            DoEvents
                        End With
                    Else
                        If rst3.Fields("Payment Date") > PreviousMonthEnd Then
                            With rst3
                                .Edit
                                .Fields("PaymentStillAvailable") = True
                                .Update
                            End With
                        End If
                    End If
            ElseIf Left(Inputdata, 17) = "Receipt Currency:" Then
                CurrencyID = Trim(Mid(Inputdata, 18, 4))
                C2 = 1
                While Currencies(C2, 1) <> CurrencyID
                    C2 = C2 + 1
                Wend
            End If
            Line Input #1, Inputdata
            Count = Count + Len(Inputdata) + 2
            If Etichetta99.Caption <> Round(Count * 100 / LFile) & "%" Then
                mPB_O2.Progressione Round(Count * 100 / LFile)
                Etichetta99.Caption = Round(Count * 100 / LFile) & "%"
                DoEvents
            End If
        Wend
        Me.Repaint
        DoEvents
        rst.Close
        rst2.Close
        rst3.Close
        Close #1
        Set rst = Nothing
        Set rst2 = Nothing
        Set rst3 = Nothing
        CurrentDb.Execute "UPDATE Tbl_Customers SET Tbl_Customers.ToSendStatement = No;"
        CurrentDb.Execute "UPDATE Tbl_Customers SET Tbl_Customers.ToSendStatement = True WHERE (DMax('[Payment Date]','Tbl_CashCollected','CustomerID=' &  [Customer_code])>[LastStatementSent]);"


        Rem put in each cash collected table the FY/FM information
        Set rst4 = CurrentDb.OpenRecordset("SELECT Tbl_CashCollected.*, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.FiscalYear FROM Tbl_CashCollected WHERE (((Tbl_CashCollected.FiscalYear) Is Null));")
        If rst4.RecordCount > 0 Then
            rst4.MoveFirst
            While Not rst4.EOF
                Set rst2 = CurrentDb.OpenRecordset("SELECT TOP 1 Tbl_MonthEnd.LatestPaymentDate, Tbl_MonthEnd.FiscalYear, Tbl_MonthEnd.FiscalQuarter, Tbl_MonthEnd.FiscalMonth FROM Tbl_MonthEnd WHERE (((Tbl_MonthEnd.MonthEnd)>=#" & Format(rst4.Fields("Tbl_CashCollected.Payment Date"), "mm/dd/yy") & "#)) ;")
                With rst4
                    .Edit
                    .Fields("Tbl_CashCollected.FiscalYear") = rst2.Fields("FiscalYear")
                    .Fields("FiscalMonth") = rst2.Fields("FiscalMonth")
                    .Fields("FiscalQuarter") = rst2.Fields("FiscalQuarter")
                    .Update
                End With
                rst4.MoveNext
            Wend
        End If
        DoEvents

        Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Customers.Customer_code, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.PaymentStillAvailable, Tbl_CashCollected.RETnumber, Tbl_Customers.Name FROM Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID WHERE (((Tbl_CashCollected.[Payment Date])>#" & Format(PreviousMonthEnd, "mm/dd/yy") & "#) AND ((Tbl_CashCollected.PaymentStillAvailable)=False));")
        If rst2.RecordCount > 0 Then
            provv = ""
            rst2.MoveFirst
            While Not rst2.EOF
                provv = provv & Chr(13) & rst2.Fields("Name") & " - Customer ID(" & rst2.Fields("Customer_code") & ") payment ID " & rst2.Fields("RETnumber")
                rst2.MoveNext
            Wend
             Call SendEmails(GetAllUserEmails, "", "Access DB - Payments have disappeared from the Cash collected report", "The following payments are disappeared from the cash report. Please check and if the case cancel the payment from the Access archive too." & Chr(13) & provv, "", "") ' if any payment disappeared from the cash collected report sends out an email to all users advising them to check.
        End If

        Rem send email with cash target report
        If rst2.RecordCount > 0 Then
            provv = "Please note that one ore more payments are disappeared from the cash collected report just uploaded."
        Else
            provv = ""
        End If

        'check if it is time to send cash target report
        Schedule = Nz(DLookup("CashTargetReportSchedule", "tblGeneral"), 0)
        If Schedule <> 0 Then
            TargetReportToBeSent = False
            ProgressiveDay = Nz(DLookup("LastCashTargetSent", "tblGeneral"), #1/1/2011#)
            If ProgressiveDay <> Date Then
                ProgressiveDay = DateAdd("d", 1, ProgressiveDay)
                Do Until (ProgressiveDay > Date) Or (TargetReportToBeSent = True)
                    DayOfTheWeek = Weekday(ProgressiveDay, vbMonday)
                    Select Case DayOfTheWeek
                        Case 1
                            TargetReportToBeSent = Schedule And 2
                        Case 2
                            TargetReportToBeSent = Schedule And 4
                        Case 3
                            TargetReportToBeSent = Schedule And 8
                        Case 4
                            TargetReportToBeSent = Schedule And 16
                        Case 5
                            TargetReportToBeSent = Schedule And 32
                    End Select
                    ProgressiveDay = DateAdd("d", 1, ProgressiveDay)
                Loop
            Else
                TargetReportToBeSent = True
            End If

        End If

        If (DLookup("ToBeSentCashTargetTo", "TblGeneral") <> "") And (TargetReportToBeSent = True) Then
            If MsgBox("Do you want to send the Cash Target email ? " & provv, vbYesNo, "") = vbYes Then

                Set rst = New ADODB.Recordset
                With rst
                    .ActiveConnection = CurrentProject.Connection
                    .Open "tblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
                    .Fields.item("LastCashTargetSent") = Date
                    .Save
                    .Close
                End With
                Set rst = Nothing


                Label1.Caption = "Running Cash Report 1/4"
                DoEvents
                Inputdata = FillGeneralCashTargetByEmail()
                Label1.Caption = "Running Cash Report 2/4"
                DoEvents
                CashTargetWithChannelByEmail = FillCashTargetWithChannelByEmail()
                Label1.Caption = "Running Cash Report 3/4"
                DoEvents
                CashTargetWithChannelByCurrency = FillCashTargetWithCurrencyByEmail()

                Label1.Caption = "Running Cash Report 4/4"
                DoEvents
                GeneralCashTargetInUSD = FillGeneralCashTargetByEmailInUSD
                '-------------------
                If Inputdata <> "" Then
                    Label1.Caption = "Sending email"
                    DoEvents
                    Rem "Cash Collected Report as of " & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), " dd mmm yyyy")
                    ConsolidatedCashTargetReport = MergeAllCashTargetRerports(Inputdata, CashTargetWithChannelByEmail, CashTargetWithChannelByCurrency, GeneralCashTargetInUSD)
                    'Call SendEmails(DLookup("ToBeSentCashTargetTo", "tblgeneral"), "", Me.Tbl_Templates.Form.Controls.Item("testo1").value, Me.Tbl_Templates.Form.Controls.Item("testo3").value, ConsolidatedCashTargetReport)
                    EmailDistributionList2 = Trim(DLookup("ToBeSentCashTargetTo", "tblgeneral"))
                    If DateDiff("d", DMax("[Payment Date]", "[Tbl_CashCollected]"), (DMin("[MonthEnd]", "[Tbl_Monthend]", "[MonthEnd]>= #" & Format(DMax("[Payment Date]", "[Tbl_CashCollected]"), "mm/dd/yy") & "#"))) <= DLookup("[CashCollectedSecondGroupDays]", "[TblGeneral]") Then
                        provv = DLookup("ToBeSentCashTargetToSecondGroup", "tblgeneral")
                        EmailDistributionList2 = EmailDistributionList2 & provv
                    End If
                    Me.Tbl_Templates.Form.Filter = "step=100"
                    Me.Tbl_Templates.Form.FilterOn = True
                    Call SendEmailsHTML(EmailDistributionList2, "", Me.Tbl_Templates.Form.Controls.item("testo1").value, Me.Tbl_Templates.Form.Controls.item("testo3").value, DLookup("PathImages", "tblgeneral") & "MergedCashTargetReports.BMP", DLookup("PathExcelDirectory", "tblgeneral") & "MergedCashTargetReports.XLS")
                    Me.Tbl_Templates.Form.FilterOn = False

                End If
            End If
        End If
        Enddate = Now()
        Label11.Caption = "End Time: " & Format(Enddate, "dd-mmm-yy   hh:mm:ss")
        Label12.Caption = "Elapsed time: " & Format(Format(Enddate - Startdate), "hh:mm:ss")
        DoEvents


        MsgBox "Import Cash Collected completed.", vbExclamation
        DoCmd.Close acForm, "MskCashImport", acSaveNo

        ''a = ExportTextDelimited("QuerycashCSV", "|")
    End If

End Sub
