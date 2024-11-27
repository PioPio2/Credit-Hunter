Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Check30_DblClick(Cancel As Integer)
    If Me.Check30.value = False Then ' if there are no attachments yet
        Call AttachDocumentsToInvoices(Me)
    Else 'if there are already attachments
        Call ShowAttachDocumentsToInvoices(Me)
    End If
End Sub

Private Sub Check30_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim a As Integer
    If Button = 2 Then
        a = CallMenu("1", Check30.value = True)
    End If
End Sub

Private Sub Form_Current()
    Me.Form.RecordsetClone.FindFirst "not isnull(CustomsInvoiceNumber)"
    Me.Controls.item("text32").ColumnHidden = Me.Form.RecordsetClone.NoMatch

    Me.Form.RecordsetClone.FindFirst "not isnull([PullTicketN#])"
    Me.Controls.item("text34").ColumnHidden = Me.Form.RecordsetClone.NoMatch
    Call TotalCalculation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim S As String
'    s = Forms.Item("MskScheduler3").Controls("maschera1").Controls("TabCtl87").Pages(Forms.Item("MskScheduler3").Controls("maschera1").Controls("TabCtl87").value).Caption
 '   i = SysCmd(acSysCmdSetStatus, "n# documents selected: " & Me.SelHeight & ", total amount: " & s & " " & Format(TotalInvoicesSelected(Me.RecordsetClone, Me.SelTop, Me.SelHeight), "##,##0.00"))
End Sub

Private Sub Testo20_AfterUpdate()
    QueryFileTochange = True
    If Testo20.Column(3) = True Then
        ChargebackFileTochange = True
    End If
End Sub

Private Sub Testo24_AfterUpdate()
    QueryFileTochange = True
    If Testo20.Column(3) = True Then
        ChargebackFileTochange = True
    End If
End Sub
Public Sub TotalCalculation()
Dim Total, Overdue130, Overdue3160, OverdueOver60, OverdueOver90, OverdueMonthEnd As Currency
Dim Monthend As Date

Me.RecordsetClone.MoveFirst
Total = 0
Overdue130 = 0
Overdue3160 = 0
OverdueOver60 = 0
OverdueOver90 = 0
OverdueMonthEnd = 0
Monthend = GetNextMonthEnd
While Not Me.RecordsetClone.EOF
    If Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") <= Now() Then
        Total = Total + Me.RecordsetClone.Fields("Amount")
    End If
    If (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") <= Now()) And (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") >= DateAdd("d", -30, Now())) Then
        Overdue130 = Overdue130 + Me.RecordsetClone.Fields("Amount")
    ElseIf (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") <= ((DateAdd("d", -31, Now())))) And (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") >= ((DateAdd("d", -60, Now())))) Then
        Overdue3160 = Overdue3160 + Me.RecordsetClone.Fields("Amount")
    ElseIf (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") < ((DateAdd("d", -60, Now())))) Then
        OverdueOver60 = OverdueOver60 + Me.RecordsetClone.Fields("Amount")
    End If
    If (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") < ((DateAdd("d", -90, Now())))) Then
        OverdueOver90 = OverdueOver90 + Me.RecordsetClone.Fields("Amount")
    End If
    If (Me.RecordsetClone.Fields("Tbl_Invoices.Overdue_Date") <= (Format(Monthend, "mm/dd/yy"))) Then
        OverdueMonthEnd = OverdueMonthEnd + Me.RecordsetClone.Fields("Amount")
    End If
    Me.RecordsetClone.MoveNext
Wend

With Me.Parent.Controls
    If OverdueOver90 > 0 Then
        .item("Label19").Visible = True
        .item("Label20").Visible = True
    Else
        .item("Label19").Visible = False
        .item("Label20").Visible = False
    End If

    .item("etichetta8").value = Format(Total, "##,##0.00")
    .item("etichetta13").Caption = Format(Overdue130, "##,##0.00")
    .item("etichetta14").Caption = Format(Overdue3160, "##,##0.00")
    .item("etichetta15").Caption = Format(OverdueOver60, "##,##0.00")
    .item("label20").Caption = Format(OverdueOver90, "##,##0.00")
    .item("Etichetta17").value = Format(OverdueMonthEnd, "##,##0.00")

End With

End Sub

Private Sub Testo24_DblClick(Cancel As Integer)
    DoCmd.OpenForm "MskShowInvoiceQuery", acNormal, , "id=" & InvoiceID.value, , acDialog
End Sub
