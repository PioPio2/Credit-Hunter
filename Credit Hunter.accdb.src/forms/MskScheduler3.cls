Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Rss As DAO.Recordset
Dim EmailOldValue, MainCurrency As String
Dim FirstFiscalMonthDay As Date
Implements IMenuMessage
Private Sub Calendar7_click()
    Testo14.value = Calendar7.value & " 08:30:00"
    Testo14.SetFocus
    Calendar7.Visible = False
End Sub
Private Sub Comando137_Click()
    DoCmd.OpenForm "MskCustomers", acNormal, , , , acDialog, Testo131.value
    Me.Refresh
End Sub
Private Sub Comando21_Click()
    On Error GoTo restoremouse
    DoCmd.Hourglass True
    Dim StatementSupport As iStatementSupport
    Set StatementSupport = New clsStatementSupportLIVE

    Dim CustomerHeader As clsCustomerHeader
    Set CustomerHeader = New clsCustomerHeader

    Call CustomerHeader.Populate(Me.Recordset.Fields("Customer_code").value)

    Dim Statement As IStatement
    Set Statement = New clsStatement
    Dim Overdue As Currency
    Overdue = 0
    Dim Outstanding As Currency
    Outstanding = 0
    Dim outCurrency As String
    Call Statement.CreateStatement(StatementSupport, CustomerHeader, Outstanding, Overdue, outCurrency, False)
    DoCmd.Hourglass False
    Exit Sub
restoremouse:
    DoCmd.Hourglass False
End Sub

Private Sub Comando41_Click()
    Dim S As String
    Dim I As Integer
    Dim InvoiceNRS As DAO.Recordset
    Dim todaysdate As String
    Dim CustomerID As String
    If IsNumeric(Testo42.value) Then    'if numberic then look for customer ID
        S = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE Tbl_Customers.Customer_code='" & Testo42.value & "';"
        Me.RecordSource = S
        If Me.Recordset.RecordCount = 0 Then
            todaysdate = Format(Date, "mm/dd/yyyy")
            Set InvoiceNRS = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Document_Number ,Tbl_Invoices.Customer_ID FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & todaysdate & "#)) GROUP BY Tbl_Invoices.Document_Number, Tbl_Invoices.Customer_ID  HAVING (((Tbl_Invoices.Document_Number)='" & Testo42.value & "'));")
            If InvoiceNRS.RecordCount > 0 Then
                CustomerID = InvoiceNRS("Customer_ID")
                S = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE Tbl_Customers.Customer_code='" & CustomerID & "';"
                Me.RecordSource = S
            End If
            'InvoiceNRS = nil
        End If

    Else    'otherwise look into customer name
        S = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE Tbl_Customers.Name LIKE '*" & Testo42.value & "*';"
        Me.RecordSource = S
        If Me.Recordset.RecordCount = 0 Then    ' if nothing is found look again into customer ID
            S = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE Tbl_Customers.Customer_code='" & Testo42.value & "';"
            Me.RecordSource = S
        End If

    End If

    If Me.Recordset.RecordCount = 0 Then
        ' se proprio non si trova niente
        MsgBox "No customer found. " & Chr(13) & "Check and try again.", vbExclamation, "Attention"
        Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers;"
    Else
        Me.TabCtl48.value = 1
    End If
End Sub

Private Sub Combo154_AfterUpdate()
    Command159.Visible = True
    Call SetTemplate(Me.Combo154.Column(2, Me.Combo154.ListIndex))
    Text162.value = Me.Sottomaschera_Tbl_Link_Customer_Internal_Email_Address.Form.GetInternalEmailCC(Tbl_Templates.Form.Text14)
End Sub


Private Sub Command125_Click()
    Dim I As Integer
    Call Form_AfterUpdate
    If Len(Testo10.value) > 0 Then
        Me.Sottomaschera_TblNotes.Form.AllowAdditions = True
        Me.Sottomaschera_TblNotes.SetFocus
        Me.Sottomaschera_TblNotes!Note.SetFocus
        DoCmd.GoToRecord , , acNewRec
        Me.Sottomaschera_TblNotes!Note = Format(Now(), "dd mmm yy - hh:mm") & vbCrLf & Testo10.value
        Me.Sottomaschera_TblNotes!CustomerCode = Testo1.value
        Me.Sottomaschera_TblNotes.Form.AllowAdditions = False
        Me.Testo10.SetFocus
        Me.Testo10.value = ""
        Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY Tbl_Customers.Index DESC;"
    End If

    Me.Recordset.Requery
    Me.SubmaskSchedulerOverview.Requery

    If Me.Recordset.RecordCount = 0 Then
        ' se proprio non si trova niente
        Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ")) ORDER BY Tbl_Customers.Index DESC;"
    Else
        Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY Tbl_Customers.Index DESC;"
    End If

    Call Form_AfterUpdate

    Me.Testo10.value = ""
    Me.SubmaskSchedulerOverview.Requery
End Sub

Private Sub Command159_Click()
    Dim rs1 As Recordset
    Dim RsBanks As Recordset
    Dim RS As DAO.Recordset
    Dim a As String
    Dim Result As Integer
    a = ""
    Result = vbYes
    If Me.Combo154.Column(6, Me.Combo154.ListIndex) = True Then
        Rem if email has to be sent
        Result = MsgBox("Do you want to send and email ?", vbYesNo, "Warning")
        Rem ask confirmation email to be sent.
        If Result = vbYes Then
            If Me.Combo154.Column(5, Me.Combo154.ListIndex) = True Then
                Rem if statement has to be included in attachment
                Set rs1 = New Recordset
                With rs1
                    .ActiveConnection = CurrentProject.Connection
                    .Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable
                    .MoveFirst
                    .Find ("Customer_code=" & Testo1.value)
                End With

                Set RsBanks = New Recordset
                With RsBanks
                    .ActiveConnection = CurrentProject.Connection
                    .Open "Tbl_banks", , adOpenKeyset, adLockOptimistic, adCmdTable
                End With

                Set RS = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.*, Tbl_Types.Descripition, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date FROM Tbl_Types INNER JOIN Tbl_Invoices ON Tbl_Types.ID = Tbl_Invoices.Type WHERE (((Tbl_Invoices.Customer_ID)=" & Testo1.value & ") AND ((Tbl_Invoices.Update_date)=#" & Format(Now(), "mm/dd/yyyy") & "#)) ORDER BY Tbl_Invoices.Currency,Tbl_Invoices.Overdue_Date,Tbl_Invoices.Document_Number;")
                a = ExcelStatement(rs1, Me.TabCtl48.Pages.item("Sheet1").Controls.item("Maschera1"), RsBanks, , True)
            End If
            Call SendEmails(Me.Recordset.Fields("email"), (Me.Recordset.Fields("ccEmail") & "," & Text162.value), Me.Testo156.value, Me.Testo155.value, a)
        End If
    End If
    Testo99.value = Me.Combo154.Column(0, Me.Combo154.ListIndex)
    If Result = vbYes Then
        Call AddNote(Testo1.value, CStr(Format(Now(), "dd mmm yy - hh:mm") & vbCrLf & "Modify status into: " & UCase(Me.Combo154.Column(1)) & vbCrLf & vbCrLf & Testo155.value))
    Else
        Call AddNote(Testo1.value, CStr(Format(Now(), "dd mmm yy - hh:mm") & vbCrLf & "Modify status into: " & UCase(Me.Combo154.Column(1)) & vbCrLf & vbCrLf & "***** THIS EMAIL WASN'T SENT BECAUSE OF OPERATOR'S REQUEST *****" & vbCrLf & Testo155.value))
    End If
    Text109.value = Now()
    MsgBox "Don't forget to update Oracle if necessary", , "Warning"
End Sub

Private Sub Command98_Enter()
    Calendar7.value = Date
    Calendar7.Visible = True
End Sub


Private Sub Form_Activate()
    Application.Echo False
    Call FindCustomerLastDate
    DoCmd.Maximize
    Command159.Visible = False
    Me.SubmaskSchedulerOverview.Requery
    '   Me.SubMaskSchedulerOverview.Form.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, Tbl_Customer_Status.ID, Tbl_Customer_Status.Description, Tbl_Areas.Area, Tbl_Areas.ID, QueryARExposureInMainCurrency.ARExposure, QueryTotalOverdueOnMonthEndInMainCurrency.TotalOverdue, QueryTotalOverdueOver90InMainCurrency.TotalOverdueOver90, Tbl_Customers.Index, * FROM ((((Tbl_Customer_Status RIGHT JOIN Tbl_Customers ON Tbl_Customer_Status.ID = Tbl_Customers.Status) LEFT JOIN Tbl_Areas ON Tbl_Customers.Area = Tbl_Areas.ID) LEFT JOIN QueryARExposureInMainCurrency ON Tbl_Customers.Customer_code = QueryARExposureInMainCurrency.Customer_ID) LEFT JOIN QueryTotalOverdueOver90InMainCurrency ON Tbl_Customers.Customer_code = QueryTotalOverdueOver90InMainCurrency.Customer_ID) LEFT JOIN QueryTotalOverdueOnMonthEndInMainCurrency ON Tbl_Customers.Customer_code = QueryTotalOverdueOnMonthEndInMainCurrency.Customer_ID WHERE (((Tbl_Customers.Credit_controller) = " & Testo5.value & ") And " & _
        " ((Tbl_Customers.NextAppointment) <=" & Date & ")) ORDER BY Tbl_Customers.Index DESC;"
    FirstFiscalMonthDay = DateAdd("d", 1, DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd <#" & Format(Date, "mm/dd/yy") & "#"))
    Application.Echo True
End Sub

Private Sub Form_AfterUpdate()
    If ChargebackFileTochange = True Then
        Call UpdateChargebackFile(Me.Recordset)
    End If
    If QueryFileTochange = True Then
        Call UpdateQueryFile(Me.Recordset)
    End If

End Sub

Private Sub Form_Current()
    Dim ctl As control
    Dim I As Integer
    Dim CashTarget, CashCollected As Currency
    Dim rst As Variant
    Application.Echo False
    Command159.Visible = False
    Me.Sottomaschera_TblNotes.Form.PutOrder
    Me.Caption = "Scheduler - Last update: " & Me.Testo60.value
    If Hour(Me.Testo60.value - Now) >= 8 Then
        Me.Caption = Me.Caption & " !!!ATTENTION DATA IS GETTING OLD !!!"
    End If

    Me.Maschera1.Visible = True
    Call TabCtl48_Change
    Me.Combo154.value = Testo99.value

    Me.Controls.item("maschera1").Visible = Me.SubMaskCurrencies.Form.Recordset.RecordCount > 0
    Me.SubmaskSchedulerOverview.Requery
    ''Call UpdateCustomerCashTarget
    '   Me.Controls("maschera1").Form.aaa
    Rem  Me.SubMaskCurrencies.Requery
    '   Me.SubMaskSchedulerOverview.Form.RecordSource = Me.Form.RecordSource
    '    Me.SubMaskSchedulerOverview.Form.Requery
    Application.Echo True
End Sub
'Public Sub UpdateCustomerCashTarget()
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim I As Integer
    Application.Echo False

    Me.PageOverview.Visible = True

    Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY Tbl_Customers.Index DESC;"
    If Me.Recordset.RecordCount = 0 Then
        Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE ((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ")  ORDER BY Tbl_Customers.Index DESC;"
    End If

    DoCmd.SelectObject acForm, Me.Name, False
    DoCmd.Maximize
    Me.Sottomaschera_TblNotes.Form.AllowAdditions = False
    QueryFileTochange = False
    ChargebackFileTochange = False
    NextMonthEnd = GetNextMonthEnd

    If QueryFileTochange = True Then
        Call UpdateQueryFile(Me.Recordset)
    End If

    Me.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY Tbl_Customers.Index DESC;"
    Combo154.ControlSource = "Status"
    Call FindCustomerLastDate
    Me.Label176.Caption = "Scheduler Overview in " & DLookup("maincurrency", "TblGeneral") & " currency:"
    'Me.SubMaskSchedulerOverview.Form.RecordSource = "SELECT Tbl_Customers.Name, Tbl_Customers.*, Tbl_Customer_Status.ID, Tbl_Customer_Status.Description, Tbl_Areas.Area, Tbl_Areas.ID, QueryARExposureInMainCurrency.ARExposure, QueryTotalOverdueOnMonthEndInMainCurrency.TotalOverdue, QueryTotalOverdueOver90InMainCurrency.TotalOverdueOver90, Tbl_Customers.Index, * FROM ((((Tbl_Customer_Status RIGHT JOIN Tbl_Customers ON Tbl_Customer_Status.ID = Tbl_Customers.Status) LEFT JOIN Tbl_Areas ON Tbl_Customers.Area = Tbl_Areas.ID) LEFT JOIN QueryARExposureInMainCurrency ON Tbl_Customers.Customer_code = QueryARExposureInMainCurrency.Customer_ID) LEFT JOIN QueryTotalOverdueOver90InMainCurrency ON Tbl_Customers.Customer_code = QueryTotalOverdueOver90InMainCurrency.Customer_ID) LEFT JOIN QueryTotalOverdueOnMonthEndInMainCurrency ON Tbl_Customers.Customer_code = QueryTotalOverdueOnMonthEndInMainCurrency.Customer_ID " & _
     " WHERE (((Tbl_Customers.Credit_controller) = " & Testo5.value & ") And ((Tbl_Customers.NextAppointment) <= format(Date(),'mm/dd/yyyy'))) ORDER BY Tbl_Customers.Index DESC;"
    Me.SubmaskSchedulerOverview.Requery
    MainCurrency = DLookup("MainCurrency", "TblGeneral")



    'Me.Controls("maschera1").Form.aaa

    Application.Echo True
End Sub

Private Sub SetTemplate(StepNumber As Integer)
    Dim I As Integer
    Me.Tbl_Templates.Form.FilterOn = False
    Me.Tbl_Templates.Form.Filter = "Step = " & StepNumber
    Me.Tbl_Templates.Form.FilterOn = True

    I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo1").value)
    Me.Testo156.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo1").value, I, Me.Recordset)

    I = TemplateBitsCount(Me.Tbl_Templates.Form.Controls("testo3").value)
    ' Me.Testo155.value = DivideTemplateInBits(Me.Tbl_Templates.Form.Controls("testo3").value, i, Me.Recordset)
End Sub

Private Sub TabCtl48_Change()
    Dim OldValue As Integer

    If Me.Controls.item("TabCtl48").Pages(Me.Controls.item("TabCtl48").value).Name = "Sheet3" Then
        If Not (IsNull(Me.Combo154.Column(2, Me.Combo154.ListIndex - 1))) Then
            Call SetTemplate(Me.Combo154.Column(2, Me.Combo154.ListIndex))
        Else
            Call SetTemplate(DLookup("[DefaultTemplate]", "TblGeneral"))
        End If
        EmailOldValue = IIf(Not (IsNull(Me.Testo155.value)), Me.Testo155.value, "")
        Me.Testo155.value = EmailOldValue
        Text162.value = Me.Sottomaschera_Tbl_Link_Customer_Internal_Email_Address.Form.GetInternalEmailCC(Tbl_Templates.Form.Text14)
    End If
End Sub


Private Sub AddNote(CustID As Long, Text As String)
    Dim I As Integer
    If Len(Text) > 0 Then
        Me.Sottomaschera_TblNotes.Form.AllowAdditions = True
        Me.Sottomaschera_TblNotes.SetFocus
        Me.Sottomaschera_TblNotes!Note.SetFocus
        DoCmd.GoToRecord , , acNewRec
        Me.Sottomaschera_TblNotes!Note = Text
        Me.Sottomaschera_TblNotes!CustomerCode = CustID
        Me.Sottomaschera_TblNotes.Form.AllowAdditions = False
        Me.Testo10.SetFocus
        Me.Testo10.value = ""
    End If
End Sub
Private Function IMenuMessage_Message(ByVal Index As Integer, ByVal item As Integer, Optional parameter As Variant) As Long
'  QUESTA E' LA FUNZIONE CHE GESTISCE IL CALL_BACK DEI MENU'
'  PERTANTO SARA' DA IMPLEMENTARE LA GESTIONE SELECT CASE idxMenu

    Select Case item
    Case 1
        Call AttachDocumentsToInvoices(Me.Maschera1.Form.ActiveControl.Form)
    Case 2
        DoCmd.OpenForm "MskDeleteInvoiceAttachments", , , "DocumentID='" & GetDocumentsToInvoices(Me.Maschera1.Form.ActiveControl.Form) & "' AND CustomerID=" & Me.Controls("Testo1") & "", , acDialog
        Dim rst2 As Variant
        Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_InvoiceAttachments.ID, Tbl_InvoiceAttachments.CustomerID, Tbl_InvoiceAttachments.DocumentID FROM Tbl_InvoiceAttachments WHERE (((Tbl_InvoiceAttachments.CustomerID)=" & Me.Controls("Testo1") & ") AND ((Tbl_InvoiceAttachments.DocumentID)='" & GetDocumentsToInvoices(Me.Maschera1.Form.ActiveControl.Form) & "'));")
        If rst2.RecordCount = 0 Then
            Me.Maschera1.Form.ActiveControl.Form.Controls("Check30") = 0
            Me.Maschera1.Form.ActiveControl.Form.Requery
        End If
    Case 3
        Call ShowAttachDocumentsToInvoices(Me.Maschera1.Form.ActiveControl.Form)
    End Select
End Function


Private Sub Testo14_LostFocus()
    Me.Requery
End Sub
