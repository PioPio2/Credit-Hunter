Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim NumOfCurrencies As DAO.Recordset
Public Function StatementExists() As Boolean
    StatementExists = Me.Sottomaschera_Tbl_Invoices.Form.Recordset.RecordCount > 0
End Function


Private Sub PopulateTabs()
    Dim ctl As control
    Dim I As Integer

    Application.Echo False
    Me.Recordset.MoveFirst
    If NumOfCurrencies.RecordCount > 0 Then
        NumOfCurrencies.MoveFirst
        For Each ctl In Me.Controls
            If ctl.ControlType = acPage Then
                Me.Controls.item(ctl.Name).Visible = Not (NumOfCurrencies.EOF)
                If NumOfCurrencies.EOF = False Then
                    If Me.Controls.item(ctl.Name).Visible = True Then
                        Me.Controls.item(ctl.Name).Caption = NumOfCurrencies.Fields("currency")
                        Select Case Me.Controls.item(ctl.Name).Name
                        Case "Pagina88"
                            'CurrentDb.QueryDefs("query1").SQL = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            'Me.Controls("Sottomaschera Tbl_Invoices").Form.RecordSource = ""
                            'Me.Controls("Sottomaschera Tbl_Invoices").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            Me.Controls("Sottomaschera Tbl_Invoices").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(Now, "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Currency)='" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                        Case "Pagina89"
                            Rem Me.Controls("SubMaskTblInvoices2RelII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            Me.Controls("SubMaskTblInvoices2RelII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(Now, "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Currency)='" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                        Case "Pagina90"
                            Rem Me.Controls("SubMaskTblInvoices2RelIII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            Me.Controls("SubMaskTblInvoices2RelII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(Now, "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Currency)='" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                        Case "Pagina91"
                            Rem Me.Controls("SubMaskTblInvoices2RelIV").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            Me.Controls("SubMaskTblInvoices2RelII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(Now, "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Currency)='" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                        Case "Pagina92"
                            Rem Me.Controls("SubMaskTblInvoices2RelV").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date) = #" & Format(Date, "mm/dd/yyyy") & "#) And ((Tbl_Invoices.Currency) = '" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                            Me.Controls("SubMaskTblInvoices2RelII").Form.RecordSource = "SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date, Tbl_Types.Descripition FROM Tbl_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Invoices.Update_date)=#" & Format(Now, "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Currency)='" & NumOfCurrencies.Fields("currency") & "')) ORDER BY Tbl_Invoices.Overdue_Date;"
                        End Select
                    End If
                    NumOfCurrencies.MoveNext
                End If
            End If
        Next ctl
    Else
        Me.Visible = False
    End If
    Application.Echo True
End Sub
Private Sub Form_Current()
    Dim ctl As control
    Dim I As Integer

    Application.Echo False

    'Me.Recordset.MoveFirst
'    Set NumOfCurrencies = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Currency FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) = '" & Me.Recordset.Fields("Customer_ID") & "') And ((Tbl_Invoices.Update_date) = #" & Format(Now, "mm/dd/yy") & "#)) GROUP BY Tbl_Invoices.Currency;")
 '   If NumOfCurrencies.RecordCount > 0 Then
  '      Call PopulateTabs
   ' Else
    '    Me.visible = False
'    End If
 '   Call Totals
    Application.Echo True
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim MainCurrency As String

    MainCurrency = DLookup("maincurrency", "TblGeneral")
    label34text = "Total collected since " & Format(FirstFiscalMonthDay, "dd mmm yy")
    'Me.Label28.Caption = "Insurance Credit Limit (" & MainCurrency & ")"
    'Me.Label22.Caption = "Internal (Oracle)Credit Limit (" & MainCurrency & ")"
    'Me.Label24.Caption = "AR bal+to be inv.+Scheduled 5 days (" & MainCurrency & ")"
    'Me.Label26.Caption = "Credit Limit available (" & MainCurrency & ")"
End Sub

Private Sub TabCtl87_Change()
    Call PopulateTabs
    Call Totals
End Sub
Private Sub Totals()
Dim CurrencyName As String
Dim CustomerID As String
Dim I, TotalToBeInvoiced, Sheduled  As Variant
Dim CustomerRecordsource As Variant
Dim TabName As String


    CurrencyName = Me.Controls.item("TabCtl87").Pages(Me.Controls.item("TabCtl87").value).Caption
    CustomerID = Me.Recordset.Fields("Customer_ID")
    TabName = Me.Parent.Form.Controls("TabCtl48").Pages("Sheet1").Controls("Maschera1").Controls("TabCtl87").Pages(Me.Controls.item("TabCtl87").value).Name
    On Error Resume Next
        Call Forms(Me.Parent.Form.Name).Form.Controls("TabCtl48").Pages("Sheet1").Controls("Maschera1").Controls("TabCtl87").Pages(TabName).Controls(0).Form.TotalCalculation
    On Error GoTo 0

    Rem insurance CL
    'Me.Label27.Caption = Format(Nz(DSum("TotalInsurance", "Tbl_Customers", "Customer_code='" & CustomerID & "'"), 0), "##,##0.00")
    'Me.Label21.Caption = Format(Nz(DSum("CreditLimit", "Tbl_CL", "Customer_code='" & CustomerID & "'"), 0), "##,##0.00")

    I = Nz(DSum("OpenARBalance", "Tbl_CL", "Customer_code='" & CustomerID & "'"), 0)
    TotalToBeInvoiced = Nz(DSum("AwaitingInvoicing", "Tbl_CL", "Customer_code='" & CustomerID & "'"), 0)
    Sheduled = Nz(DSum("AmtScheduledTom", "Tbl_CL", "Customer_code='" & CustomerID & "'"), 0)
    'Me.Label23.Caption = Format(i + TotalToBeInvoiced + Sheduled, "##,##0.00")

    I = Nz(DSum("CreditLimit", "Tbl_CL", "Customer_code='" & CustomerID & "'") - I - TotalToBeInvoiced - Sheduled, 0)
'    If i < 0 Then
 '       Me.Label25.BackColor = RGB(255, 0, 0)
  '      Me.Label25.BackStyle = 1
   '     Me.Label26.BackColor = RGB(255, 0, 0)
    '    Me.Label26.BackStyle = 1
'    Else
 '       Me.Label25.ForeColor = RGB(0, 0, 0)
  '      Me.Label25.BackStyle = 0
   '     Me.Label26.ForeColor = RGB(0, 0, 0)
    '    Me.Label26.BackStyle = 0
'    End If

'    Me.Label25.Caption = Format(i, "##,##0.00")

'    Me.Label29.Caption = Format(DLookup("HighestExposure", "Tbl_Customers", "Customer_code='" & CustomerID & "'"), "##,##0.00")
    Me.Etichetta12.Caption = "Total overdue as of today (" & CurrencyName & ")"
    Me.Etichetta9.Caption = "Overdue 1-30 days as of today (" & CurrencyName & ")"
    Me.Etichetta10.Caption = "Overdue 31-60 days as of today (" & CurrencyName & ")"
    Me.Etichetta11.Caption = "Overdue over 60 as of today (" & CurrencyName & ")"
    Me.Etichetta16.Caption = "Total overdue on month end (" & CurrencyName & ")"
    Me.Label19.Caption = "Overdue over 90 days (check insurance obligations) (" & CurrencyName & ")"
End Sub


Public Sub aaa()
Dim CashTarget, CashCollected As Currency
On Error GoTo ExitSub
If Me.Parent.Controls("submaskcurrencies").Form.Recordset.RecordCount > 0 Then
        If (Me.Parent.Controls("submaskcurrencies").Controls("Currency").value <> MainCurrency) Then
            CashTarget = Nz(Me.Parent.Recordset.Fields("MonthlyTargetInMainCurrency") / DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='" & Me.Parent.Controls("submaskcurrencies").Controls("Currency").value & "'"), 0)
            FirstFiscalMonthDay = DateAdd("d", 1, DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd <#" & Format(Date, "mm/dd/yy") & "#"))
            Set rst = CurrentDb.OpenRecordset("SELECT Sum([Original amount]) AS AmountInEUR FROM Tbl_CashCollected INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID WHERE  (((Tbl_CashCollected.CustomerID)=" & Me.Form.Recordset.Fields("Customer_ID") & ") AND ((Tbl_CashCollected.[Payment Date]) >=#" & Format(FirstFiscalMonthDay, "mm/dd/yy") & "#));")
            CashCollected = Nz(rst.Fields("AmountInEUR"), 0)
        Else
            CashTarget = Nz(Me.Recordset.Fields("MonthlyTargetInMainCurrency"), 0)
            FirstFiscalMonthDay = DateAdd("d", 1, DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd <#" & Format(Date, "mm/dd/yy") & "#"))
            Set rst = CurrentDb.OpenRecordset("SELECT Sum([ExchangeRate]*[Original amount]) AS AmountInEUR FROM Tbl_CashCollected INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID WHERE  (((Tbl_CashCollected.CustomerID)=" & Me.Form.Recordset.Fields("Customer_code") & ") AND ((Tbl_CashCollected.[Payment Date]) >=#" & Format(FirstFiscalMonthDay, "mm/dd/yy") & "#));")
            CashCollected = Nz(rst.Fields("AmountInEUR"), 0)
        End If
        Me.Controls("label32").Caption = "Customer cash target in " & Me.Parent.Controls("submaskcurrencies").Controls("Currency").value
        Me.Controls("text33") = Format(CashTarget, "##,##0.00")
        label34text = "Total collected since " & Format(FirstFiscalMonthDay, "dd mmm yy")
        Me.Controls("label34").Caption = label34text & " in " & Me.Parent.Controls("submaskcurrencies").Controls("Currency").value
        Me.Controls("text35") = Format(CashCollected, "##,##0.00")
        Me.Controls("label36").Caption = "Still to be collected in " & Me.Parent.Controls("submaskcurrencies").Controls("Currency").value
        Me.Controls("text37").value = Format(CashTarget - CashCollected, "##,##0.00")
        If CashTarget <> 0 Then
            Me.Controls("text39").value = CInt((CashTarget - CashCollected) * 100 / CashTarget) & "%"
        Else
            Me.Controls("text39").value = "0%"
        End If
    End If
 Exit Sub
ExitSub:
On Error GoTo 0
 Exit Sub
End Sub
