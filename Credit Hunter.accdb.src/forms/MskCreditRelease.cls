Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Private mPB_O2 As clsPROGRESS
Dim CountFile, LFile As Long
Dim MainCurrency As String
Private Sub Comando0_Click()
Dim rst As Recordset
Dim Inputdata, StrInputFileName, StrFilter As String
    StrFilter = ahtAddFilterItem(StrFilter, "Text files(*.txt)", "*.txt")
    StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
        DialogTitle:="Please select an input file...", _
        Flags:=ahtOFN_HIDEREADONLY)
    If StrInputFileName <> "" Then
        Me.Etichetta99.Visible = True
        Me.Repaint
        CountFile = 0
        LFile = FileLen(StrInputFileName)
        Rem mPB_O2.DefinisciPB
        Open StrInputFileName For Input As #1    ' Apre il file per l'input.
        While InStr(1, Inputdata, "Schedule Date Threshold") = 0
            Line Input #1, Inputdata
            CountFile = CountFile + Len(Inputdata) + 2
        Wend
        Close #1
        If CDate(Mid(Inputdata, 63, 10)) <> DateAdd("d", Date, 7) Then
            MsgBox "Please check report date because it's different than today+5 days. I will proceed with the upload anyway now.", vbExclamation
        End If

        Rem Insert Horizon Date Limit
        Set rst = New Recordset
        With rst
            .ActiveConnection = CurrentProject.Connection
            .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
            .Fields("CLHorizonDateLimit") = CDate(Mid(Inputdata, 63, 10))
            .Update
            .Close
        End With
        Set rst = Nothing

        Call InsertNewCreditLimit(1, StrInputFileName)
        Me.Testo122.enabled = False
        Me.Testo123.enabled = False
        Me.Refresh
        Me.Requery
        Me.Testo122.enabled = True
        Me.Testo123.enabled = True
        Call LabelPaint

        Rem #########  Updates Historical Credit Limits #######
        Call UpdateHistoricalCL

        Rem ######### SEND CL AUTOMATICALLY #########
        If Not (IsNull(DLookup("ToBeSentCLto", "TblGeneral"))) Then
            Call CLLimitReport
            Call SendEmails(DLookup("ToBeSentCLto", "TblGeneral"), "", "Updated Credit Limits", "", GetPathExcelDirectory() & "Updated CL Report.xls")
        End If
    End If
End Sub

Private Sub InsertNewCreditLimit(I As Integer, nfile As Variant)
Dim Country, Inputdata, InputData2, CustCode, AreaName As String
Dim rst As Recordset
Dim RstCustomers As Recordset
Dim RstAreas As Recordset
Dim RstCountries As Recordset
Dim AreaCode As Integer

If nfile <> "" Then
    DoCmd.OpenQuery ("QueryToReleaseOrder")
    DoCmd.OpenQuery ("QueryToSendReleaseRequest")

    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CurrentProject.Connection
    rst.Open "Tbl_CL", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set RstCustomers = New ADODB.Recordset
    RstCustomers.ActiveConnection = CurrentProject.Connection
    RstCustomers.Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set RstAreas = New ADODB.Recordset
    RstAreas.ActiveConnection = CurrentProject.Connection
    RstAreas.Open "Tbl_Areas", , adOpenKeyset, adLockOptimistic, adCmdTable

    Open nfile For Input As #1    ' Apre il file per l'input.
    Open nfile For Input As #2    ' Apre copia del file per l'input.
    Set mPB_O2 = New clsPROGRESS
    mPB_O2.CreaPB Me.shPB_O2, Orizzontale
    mPB_O2.DefinisciPB
    Me.TabCtl57.Visible = False
    Me.Etichetta99.Visible = True
    Line Input #2, InputData2
    While Not EOF(1)
        Line Input #1, Inputdata
        If Not EOF(2) Then
            Line Input #2, InputData2
        End If
        CountFile = CountFile + Len(Inputdata) + 2
        If IsNumeric(Trim(Mid$(Inputdata, 45, 5))) Then
            With rst
                AreaName = Trim(Left(Inputdata, 12))
                If Not IsNumeric(Trim(Mid$(InputData2, 45, 5))) And (Asc(Left(InputData2, 1)) > 12) Then
                    AreaName = Trim(AreaName & " " & Trim(Left(InputData2, 12)))
                End If
                If Not RstAreas.EOF Then
                    RstAreas.MoveFirst
                    RstAreas.Find ("Area='" & AreaName & "'")
                End If
                If RstAreas.EOF Then
                    RstAreas.AddNew
                    RstAreas.Fields("Area") = AreaName
                    RstAreas.Update
                End If

                AreaCode = RstAreas.Fields("ID")

                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find ("Customer_code='" & Trim(Mid$(Inputdata, 45, 5)) & "'")
                End If
                If .EOF Then
                    .AddNew
                    .Fields("Customer_code") = Trim(Mid$(Inputdata, 45, 5))
                End If
                CustCode = Trim(Mid$(Inputdata, 45, 5))
                RstCustomers.MoveFirst
                RstCustomers.Find ("Customer_code='" & CustCode & "'")
                If Not RstCustomers.EOF Then
                    RstCustomers.Fields("Area") = AreaCode
                    RstCustomers.Update
                End If

                If I = 1 Then
                    .Fields("Currency") = Trim(Mid$(Inputdata, 51, 3))
                    If Trim(Mid$(Inputdata, 54, 17)) = "UNLIMITED" Then
                        .Fields("CreditLimit") = 99999999
                    Else
                        .Fields("CreditLimit") = Trim(Mid$(Inputdata, 54, 17))
                    End If
                    .Fields("OpenARBalance") = Trim(Mid$(Inputdata, 106, 14))
                    .Fields("AwaitingInvoicing") = Mid$(Inputdata, 124, 13)
                    .Fields("AmtScheduledTom") = Mid$(Inputdata, 150, 15)
                Else
                    .Fields("AmtScheduled8Dyas") = Mid$(Inputdata, 150, 15)
                End If
                mPB_O2.Progressione (CInt(100 * CountFile / LFile))
                Me.Etichetta99.Caption = (CStr(Int(100 * CountFile / LFile))) & "%"
                .Update
                Me.Repaint
            End With
        End If
    Wend
    rst.Close
    RstCustomers.Close
    RstAreas.Close
    Set rst = Nothing
    Set RstAreas = Nothing
    Set RstCustomers = Nothing
    Set mPB_O2 = Nothing
    Me.TabCtl57.Visible = True
    Me.Etichetta99.Visible = False
    Close #1
    Close #2

    Rem update import datetime
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        If I = 1 Then
            .Fields("Update_CL+1") = Now()
        Else
            .Fields("Update_CL+8") = Now()
        End If
        .Update
        .Close
    End With
    Set rst = Nothing
    MsgBox "Import credit limits completed", vbExclamation
    Me.Repaint
End If

End Sub
Private Sub Comando12_Click()
    Call InsertCustomersFail
    Me.Recordset.MoveFirst
    Call Form_Open(0)
End Sub

Private Sub Comando39_Click()
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim r, a As Recordset
Dim FN, DirSave, b As String
Dim b2 As String
Dim Ordernum As String
Dim EmailTo, EmailCC As String
Dim I As Integer
Dim LstDate As Date
Dim TotalOrder As Currency
    Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory & "CL BLANK FORM.XLS")
    ExcApp.Visible = True
    With ExcDoc.Sheets(1)
        .Cells(3, 2) = Format(Date, "DD-MMM-YYYY")
        .Cells(6, 2) = Testo17.value
        .Cells(6, 3) = Testo40.value
        .Cells(6, 4) = Me.Sottomaschera_Tbl_credit_check_failures![Sub-Region]
        .Cells(6, 5) = Testo19.value
        .Cells(6, 6) = Testo50.value
        .Cells(11, 3) = Testo41.value
        .Cells(13, 3) = Testo53.value
        .Cells(15, 3) = CreditLimit.value
        .Cells(17, 3) = OpenARBalance.value
        .Cells(19, 3) = AwaitingInvoicing.value
        .Cells(21, 3) = AmtScheduledTom.value
Rem        .Cells(27, 3) = AmtScheduled8Dyas.Value
        If Etichetta56.Visible = True Then
            TotalOrder = DSum("[amount]", "Tbl_credit_check_failures", "[customer number]= " & Testo19.value)
            .Cells(9, 5) = "=Sum(" & TotalOrder & "/" & Testo55.value & ")"
        Else
            .Cells(9, 5) = DSum("[amount]", "Tbl_credit_check_failures", "[customer number]= " & Testo19.value)
        End If
        .Cells(11, 5) = Format(DMin("[Schedule date]", "Tbl_credit_check_failures", "[customer number]= " & Testo19.value), "DD-MMM-YYYY")

        If Etichetta56.Visible = True Then
            If IsNull(Testo32.value) Then Testo32.value = 0
            .Cells(19, 5) = "=Sum(" & Testo32.value & "/" & Testo55.value & ")"
        Else
            .Cells(19, 5) = Testo32.value
        End If

        .Cells(21, 5) = Format(Testo33.value, "DD-mmm-YYYY")
        .Cells(23, 5) = Testo34.value
        .Cells(25, 5) = Testo35.value
        .Cells(36, 5) = GetNameCreditController(fOSUserName)

        .Cells(40, 2) = "Currency"
        .Cells(40, 3) = "Current"
        .Cells(40, 4) = "1-30 Days"
        .Cells(40, 5) = "31-60 Days"
        .Cells(40, 6) = "61+ Days"
        .Cells(40, 7) = "Balance "

         LstDate = FindLastDate()
        Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Currency, Sum(IIf([overdue_date]>=#" & Format(LstDate, "mm/dd/yyyy") & "# ,[Amount],0)) AS [Current], Sum(IIf([overdue_date]<#" & Format(LstDate, "mm/dd/yyyy") & "# And [overdue_date]>=#" & Format(DateAdd("d", LstDate, -30), "mm/dd/yyyy") & "#,[Amount],0) ) AS [1-30days], Sum(IIf([overdue_date]<=#" & Format(DateAdd("d", LstDate, -31), "mm/dd/yyyy") & "# And [overdue_date]>=#" & Format(DateAdd("d", LstDate, -60), "mm/dd/yyyy") & "#,[Amount],0)) AS [31-60days], Sum(IIf([overdue_date]<=#" & Format(DateAdd("d", LstDate, -61), "mm/dd/yyyy") & "#,[Amount],0)) AS [61+days] FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) =#" & Format(LstDate, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Customer_ID) = " & Testo19.value & "))GROUP BY Tbl_Invoices.Currency;")
        If r.RecordCount > 0 Then
            r.MoveFirst
            For I = 1 To r.RecordCount
            .Cells(40 + I, 2) = Format(r.Fields("currency"), "##,##0.00")
            .Cells(40 + I, 3) = Format(r.Fields("current"), "##,##0.00")
            .Cells(40 + I, 4) = Format(r.Fields("1-30days"), "##,##0.00")
            .Cells(40 + I, 5) = Format(r.Fields("31-60days"), "##,##0.00")
            .Cells(40 + I, 6) = Format(r.Fields("61+days"), "##,##0.00")
            .Cells(40 + I, 7) = Format(r.Fields("current") + r.Fields("1-30days") + r.Fields("31-60days") + r.Fields("61+days"), "##,##0.00")
            r.MoveNext
            Next I
        End If
    End With

    DirSave = "C:\Users\" & fOSUserName() & "\Releases\"
    If Dir(DirSave, 16) = "" Then
        MkDir (DirSave)
    End If

    DirSave = "C:\Users\" & fOSUserName() & "\Releases\" & Testo40.value & "\"
    If Dir(DirSave, 16) = "" Then
        MkDir (DirSave)
    End If

    DirSave = "C:\Users\" & fOSUserName() & "\Releases\" & Testo40.value & "\" & Format((Date), "dd-mm-yyyy") & "\"
    If Dir(DirSave, 16) = "" Then
        MkDir (DirSave)
    End If

    FN = NormalizeFileName(Testo17.value)
    LstDate = Now()
    FN = DirSave & FN & " - " & Format((LstDate), "dd mmm yyyy - hh.mm.ss") & ".xls"
    ExcApp.ActiveWorkbook.SaveAs FileName:=FN, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False

    ExcApp.Quit
    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    b = FN

Rem ### inizio elenco linee da rilasciare ###
    Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory & "Lines released.xls")
    ExcApp.Visible = True
    I = 2
    Ordernum = ""
    Me.Sottomaschera_Tbl_credit_check_failures.Form.RecordsetClone.MoveFirst
    With Me.Sottomaschera_Tbl_credit_check_failures.Form.RecordsetClone
        While Not Me.Sottomaschera_Tbl_credit_check_failures.Form.RecordsetClone.EOF
            ExcDoc.Sheets(1).Cells(I, 1) = .Fields("hold type")
            ExcDoc.Sheets(1).Cells(I, 2) = .Fields("Hold Name")
            ExcDoc.Sheets(1).Cells(I, 3) = .Fields("Date Hold Applied")
            ExcDoc.Sheets(1).Cells(I, 4) = .Fields("Hold Until Date")
            ExcDoc.Sheets(1).Cells(I, 5) = .Fields("Hold Comments")
            ExcDoc.Sheets(1).Cells(I, 6) = .Fields("Sub-Region")
            ExcDoc.Sheets(1).Cells(I, 7) = .Fields("Country")
            ExcDoc.Sheets(1).Cells(I, 8) = .Fields("customer Name")
            ExcDoc.Sheets(1).Cells(I, 9) = .Fields("customer number")
            ExcDoc.Sheets(1).Cells(I, 10) = .Fields("Account Specialist")
            ExcDoc.Sheets(1).Cells(I, 11) = .Fields("Logitech Item Number")
            ExcDoc.Sheets(1).Cells(I, 12) = .Fields("List Price")
            ExcDoc.Sheets(1).Cells(I, 13) = .Fields("Requested Quantity")
            ExcDoc.Sheets(1).Cells(I, 14) = .Fields("Currency Code")
            ExcDoc.Sheets(1).Cells(I, 15) = .Fields("Amount")
            ExcDoc.Sheets(1).Cells(I, 16) = .Fields("Order number")
            If Ordernum <> .Fields("Order number") Then
                Ordernum = .Fields("Order number")
                ExcDoc.Sheets(1).Cells(I, 16).Font.Bold = True
            End If
            ExcDoc.Sheets(1).Cells(I, 17) = .Fields("Order Line Number")
            ExcDoc.Sheets(1).Cells(I, 18) = .Fields("Order Date")
            ExcDoc.Sheets(1).Cells(I, 19) = .Fields("Requested Date")
            ExcDoc.Sheets(1).Cells(I, 20) = .Fields("Schedule Date")
            ExcDoc.Sheets(1).Cells(I, 21) = .Fields("Active Hold")
            ExcDoc.Sheets(1).Cells(I, 22) = .Fields("Open Line")
            ExcDoc.Sheets(1).Cells(I, 23) = .Fields("Line Status")
            ExcDoc.Sheets(1).Cells(I, 24) = .Fields("Hold Criteria")
            ExcDoc.Sheets(1).Cells(I, 25) = .Fields("Tax Code")
            I = I + 1
            Me.Sottomaschera_Tbl_credit_check_failures.Form.RecordsetClone.MoveNext
        Wend
    End With


    ExcDoc.Sheets(1).Cells(I + 2, 14) = "Total:"
    ExcDoc.Sheets(1).Cells(I + 2, 14).Font.Bold = True

    ExcDoc.Sheets(1).Cells(I + 2, 15) = "=Sum(o1:o" & I + 1 & ")"
    ExcDoc.Sheets(1).Cells(I + 2, 15) = Format(ExcDoc.Sheets(1).Cells(I + 2, 15), "##,##0.00")
    ExcDoc.Sheets(1).Cells(I + 2, 15).Font.Bold = True

    DirSave = "C:\Users\" & fOSUserName() & "\Releases\" & Testo40.value & "\"
    If Dir(DirSave, 16) = "" Then
        MkDir (DirSave)
    End If

    DirSave = "C:\Users\" & fOSUserName() & "\Releases\" & Testo40.value & "\" & Format((Date), "dd-mm-yyyy") & "\"
    If Dir(DirSave, 16) = "" Then
        MkDir (DirSave)
    End If

    FN = NormalizeFileName(Testo17.value)
    FN = DirSave & FN & " - Lines to release - " & Format((LstDate), "dd mmm yyyy - hh.mm.ss") & ".xls"
    ExcApp.ActiveWorkbook.SaveAs FileName:=FN, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False

    ExcApp.Quit
    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    b2 = FN
Rem ### fine elenco linee da rilasciare ###

    EmailTo = ""
    EmailCC = ""
    Call GetApproverEmailAddress(((CCur(Etichetta104.Caption) * -1) / DLookup("exchangeRate", "tbl_currencies", "CurrencyID='USD'")), EmailTo, EmailCC)
    Call SendEmails(CStr(EmailTo), EmailCC, Me.Testo123.value, Me.Testo122.value, b, b2)
End Sub
Private Sub Command124_Click()
    CurrentDb.Execute "UPDATE Tbl_credit_check_failures SET Tbl_credit_check_failures.Released = Not [released] WHERE (((Tbl_credit_check_failures.[Customer Number])=" & Testo19.value & "));"
    Me.Sottomaschera_Tbl_credit_check_failures.Requery
End Sub

Private Sub Command125_Click()
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim S As String
Dim Rec As Variant
Dim I, a As Integer

    I = 0
    Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory & "Failed Credit Check Releases Template.XLS")
    ExcApp.Visible = True
    With ExcDoc.Sheets(1)
        S = .Cells(4, 3)
        S = Replace(S, "TO BE REPLACED", Format(Testo71.value, "dd-mmm-yyyy"))
        .Cells(4, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE ((Not (Tbl_credit_check_failures.[Hold Name])='LOGI Manual Credit Hold')); ")
        .Cells(4, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        S = .Cells(6, 3)
        S = Replace(S, "TO BE REPLACED", Format(Testo71.value, "dd-mmm-yyyy"))
        .Cells(6, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE ((Not (Tbl_credit_check_failures.[Hold Name])='LOGI Manual Credit Hold') AND ((Tbl_credit_check_failures.Released)=True));")
        .Cells(6, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR, Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.Released, Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID GROUP BY Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.Released, Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country HAVING (((Tbl_credit_check_failures.[Hold Name]) <> 'LOGI Manual Credit Hold') And ((Tbl_credit_check_failures.Released) = False)) ORDER BY Sum([Amount]*[ExchangeRate]) DESC;")
        If Rec.RecordCount > 0 Then
            Rec.MoveFirst
            I = 0
            While Not Rec.EOF
                .Cells(12 + I, 3) = Rec.Fields("country")
                .Cells(12 + I, 4) = DLookup("Name", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                .Cells(12 + I, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")
                .Cells(12 + I, 6) = DLookup("ReleaseNotes", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                I = I + 1
                Rec.MoveNext
            Wend
            .Range("A" & 12 + I & ":z100").Delete Shift:=xlUp
            I = I + 14
        End If
        If I = 0 Then I = 103
        S = .Cells(I, 3)
        S = Replace(S, "TO BE REPLACED", Format(Testo71.value, "dd-mmm-yyyy"))
        .Cells(I, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE (((Tbl_credit_check_failures.[Hold Type])='LOGI Manual Credit Hold'));")
        .Cells(I, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        I = I + 4
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR, Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country, Tbl_credit_check_failures.[Hold Type] FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID GROUP BY Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country, Tbl_credit_check_failures.[Hold Type] HAVING (((Tbl_credit_check_failures.[Hold Type])='LOGI Manual Credit Hold')) ORDER BY Sum([Amount]*[ExchangeRate]) DESC;")
        If Rec.RecordCount > 0 Then
            Rec.MoveFirst
            a = 0
            While Not Rec.EOF
                .Cells(I + a, 3) = Rec.Fields("country")
                .Cells(I + a, 4) = DLookup("Name", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                .Cells(I + a, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")
                .Cells(I + a, 6) = DLookup("ReleaseNotes", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                I = I + 1
                Rec.MoveNext
            Wend
            .Range("A" & I & ":z300").Delete Shift:=xlUp
        End If
        .Range("A1").Select
    End With
End Sub
Private Sub Form_Activate()
    DoCmd.Maximize
End Sub
Private Sub Form_Current()
    Dim I As Integer
    Dim a As String
    If Not IsNull(Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode]) Then
        Etichetta86.Caption = "Total order amount in original currency (" & Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode] & "):"
        Etichetta100.Caption = Format(((DLookup("[ExchangeRate]", "[Tbl_Currencies]", "[CurrencyID]='" & Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode] & "'")) * DSum("[amount]", "[Tbl_credit_check_failures]", "[Customer Number]=" & Testo19.value)), "##,##0.00")
        Etichetta56.Visible = Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode] <> MainCurrency
        If Etichetta56.Visible = True Then
            Etichetta56.Caption = "Exchange rate: 1 " & MainCurrency & " = ?? " & Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode]
        End If
        Testo55.Visible = Etichetta56.Visible
        Etichetta104.Caption = Format(CreditLimit - OpenARBalance - AwaitingInvoicing - (DLookup("[ExchangeRate]", "[Tbl_Currencies]", "[CurrencyID]='" & Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode] & "'") * DSum("[amount]", "[Tbl_credit_check_failures]", "[Customer Number]=" & Testo19.value)), "##,##0.00")
        Etichetta105.Caption = Format((CreditLimit - OpenARBalance - AwaitingInvoicing - (DLookup("[ExchangeRate]", "[Tbl_Currencies]", "[CurrencyID]='" & Me.Sottomaschera_Tbl_credit_check_failures![CurrencyCode] & "'")) * DSum("[amount]", "[Tbl_credit_check_failures]", "[Customer Number]=" & Testo19.value)) / DLookup("[ExchangeRate]", "[Tbl_Currencies]", "[CurrencyID]='USD'"), "##,##0.00")

        If Me.TabCtl57.value = 1 Then
            I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
            a = GetApproverEmailAddress(((CCur(Etichetta105.Caption))) * -1, EmailTo, CStr(EmailCC))
            If Me.Testo122.enabled = True Then
                I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo1").value)
                Me.Testo123.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo1").value, I, Me.Recordset)

                I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
                Me.Testo122.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo3").value, I, Me.Recordset, a)
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim I As Integer
    DoCmd.SelectObject acForm, Me.Name, False
    DoCmd.Maximize
    Me.Etichetta99.Visible = False

    Me.Tbl_Templates.Form.FilterOn = False
    Me.Tbl_Templates.Form.Filter = "Step = 99"
    Me.Tbl_Templates.Form.FilterOn = True

    I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo1").value)
    Me.Testo123.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo1").value, I, Me.Recordset)

    I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
    Me.Testo122.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo3").value, I, Me.Recordset)
    Call LabelPaint

    Me.Sottomaschera_TblNotes.Form.PutOrder
    Comando39.enabled = CDate(Testo71.value) <= CDate(Testo79.value)
    MainCurrency = DLookup("[MainCurrency]", "[TblGeneral]")
    Label110.Caption = "Total order amount in " & MainCurrency & " currency:"
    Command124.Visible = DLookup("superuser", "tbl_users", "UserName='" & fOSUserName() & "'") = True
    Command125.Visible = Command124.Visible
    Box127.Visible = Command124.Visible
    Text126.Visible = Command124.Visible
End Sub
Private Sub LabelPaint()
    If DateAdd("h", 8, Testo71.value) <= Now() Then
        Me.Testo71.ForeColor = RGB(255, 0, 0)
        Me.Testo71.FontBold = True
    Else
        Me.Testo71.ForeColor = RGB(0, 0, 0)
        Me.Testo71.FontBold = False
    End If
    If DateAdd("n", 60, Testo79.value) <= Now() Then
        Me.Testo79.ForeColor = RGB(255, 0, 0)
        Me.Testo79.FontBold = True
    Else
        Me.Testo79.ForeColor = RGB(0, 0, 0)
        Me.Testo79.FontBold = False
    End If
End Sub

Private Sub Form_Timer()
Static JustDoIt    As Boolean
    If JustDoIt = False Then
        JustDoIt = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mPB_O2 = Nothing
End Sub

Private Sub TabCtl57_Change()
Dim I As Integer
Dim a As String
    I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
    a = GetApproverEmailAddress(((CCur(Etichetta105.Caption))) * -1, EmailTo, CStr(EmailCC))
    If Me.Testo122.enabled = True Then
        I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo1").value)
        Me.Testo123.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo1").value, I, Me.Recordset)

        I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
        Me.Testo122.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo3").value, I, Me.Recordset, a)
    End If

    If Testo55.Visible = True Then
        If (IsNumeric(Testo55.value) = False) And (Me.TabCtl57.value <> 0) Then
            Me.TabCtl57.value = 0
            a = MsgBox("Please insert a valid exchange rate", vbCritical, "")
            Testo55.value = ""
        End If
    End If
End Sub

Private Sub Testo87_BeforeUpdate(Cancel As Integer)
    If MsgBox("The default text has been changed, is this the new default text ?", vbYesNo) = vbNo Then
        Cancel = False
        VarUpdate = False
    End If
End Sub

Private Sub InsertCustomersFail()
Dim StrInputFileName, StrFilter As String
Dim rst As Recordset
Dim Riga, I As Integer
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim Error As Boolean
Dim Header
Dim NRows As Integer

Header = Array("Hold Type", "Hold Name", "Date Hold Applied", "Hold Until Date", "Hold Comments", "Sub-Region", "Country", "Customer Name", "Customer Number", "Account Specialist", "Logitech Item Number", "Item Description", "List Price", "Requested Quantity", "Currency Code", "Amount", "Order Number", "Order Line Number", "Order Date", "Requested Date", "Schedule Date", "Active Hold", "Open Line", "Line Status", "Hold Criteria", "Tax Code")
    StrFilter = ahtAddFilterItem(StrFilter, "Excel files(*.xls)", "*.xls")
    StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select an input file...", _
                Flags:=ahtOFN_HIDEREADONLY)
    If StrInputFileName <> "" Then
        Me.TabCtl57.Visible = False
        Me.Etichetta99.Visible = True
        Me.Repaint
        Set mPB_O2 = New clsPROGRESS
        mPB_O2.CreaPB Me.shPB_O2, Orizzontale
        mPB_O2.DefinisciPB
        Me.Etichetta99.Visible = True
        NRows = NumMaxRows(CStr(StrInputFileName), "")
        DoCmd.OpenQuery ("QueryToReleaseOrder")
        DoCmd.OpenQuery ("QueryToSendReleaseRequest")
        Set rst = New ADODB.Recordset
        rst.ActiveConnection = CurrentProject.Connection
        rst.Open "Tbl_credit_check_failures", , adOpenKeyset, adLockOptimistic, adCmdTable
        Set ExcApp = CreateObject("Excel.Application")
        Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
        ExcApp.Visible = False
        Riga = 2
        Error = False

        With ExcDoc.Sheets("Credit Check Failures")
            Error = False
            For I = 0 To 25
                If Header(I) <> .Cells(1, I + 1) Then
                    Error = True
                End If
            Next I
            If Error = True Then
                GoTo uscita
            End If
            DoCmd.OpenQuery ("QueryEraseTblCreditCheckFailures")
            While .Cells(Riga, 1) <> ""
                If .Cells(Riga, 23) = "Awaiting Shipping" Or .Cells(Riga, 23) = "Picked Partial" Then
                    rst.AddNew
                    For I = 0 To 24
                        rst.Fields(Header(I)) = .Cells(Riga, I + 1)
                    Next I
                    rst.Update
                End If
                Riga = Riga + 1
                mPB_O2.Progressione (CInt(100 * Riga / NRows))
                Etichetta99.Caption = (CInt(100 * Riga / NRows)) & "%"
                Me.Repaint
            Wend
        End With

        ExcDoc.Close
        Set ExcDoc = Nothing
        Set ExcApp = Nothing
        rst.Close

        Rem update import datetime
        Set rst = New ADODB.Recordset
        With rst
            .ActiveConnection = CurrentProject.Connection
            .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
            .MoveFirst
            .Fields("Update_Customers_Failing") = Now()
            .Update
            .Close
        End With
        Set rst = Nothing

        Me.TabCtl57.Visible = True
        Me.Etichetta99.Visible = False
        Rem mPB_O2.Progressione (0)
        Set mPB_O2 = Nothing
        Me.Testo122.enabled = False
        Me.Testo123.enabled = False
        Me.Refresh
        Me.Requery
        Me.Testo122.enabled = True
        Me.Testo123.enabled = True
        MsgBox "Import customer who fail the credit check completed.", vbExclamation
        Call LabelPaint
    End If
    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    Exit Sub
uscita:
    MsgBox "Import customer who fail the credit check FAILED. The hearder in the Excel file doesn't match the standard", vbCritical
    ExcDoc.Close
    Me.TabCtl57.Visible = True
    Etichetta99.Visible = False

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    rst.Close
End Sub
