Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command4_Click()
    Dim ExcApp As Excel.Application
    Dim ExcDoc As Excel.Workbook
    Dim RS As Recordset
    Dim I, TabCashForecast, row, col, ColCustID As Integer
    Dim S, StrFilter As String
    Dim Found As Boolean
    Dim rs1 As Variant

    If CDate(Text0.value) > CDate(Text2.value) Then
        MsgBox "From date has to be before To date. Please check.", vbOKOnly + vbCritical, "Error: "
    Else
        StrFilter = ahtAddFilterItem(StrFilter, "Excel files(*.xls, *.xlsx)", "*.xls;*.xlsx")
        StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
                        DialogTitle:="Please select an input file...", _
                        Flags:=ahtOFN_HIDEREADONLY Or ahtOFN_EXPLORER)

        If StrInputFileName <> "" Then
            Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
            Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
            ExcApp.Visible = True
            Found = False
            With ExcApp
                For I = 1 To ExcApp.Sheets.Count
                    If UCase(.Sheets(I).Name) = "cash forecast" Then
                        Found = True
                        TabCashForecast = I
                        Exit For
                    End If
                Next I

                Rem if tab cash forecast exists:
                If Found Then
                    row = 0
                    col = 0
                    Found = False
                    For row = 1 To 10
                        For col = 1 To 100
                            If .Sheets(TabCashForecast).Cells(row, col) = UCase("Cust. No.") Then

                                Found = True
                                Exit For
                            End If
                            If Found Then Exit For
                        Next col
                        If Found Then Exit For
                    Next row
                    If Found Then
                        ColCustID = col
                    Else
                        MsgBox "Column Cust. No. in the Excel file selected hasn't been found. Cash forecast is NOT updated in Access", vbCritical

                    End If
                Else
                    MsgBox "Tab Cash forecast in the Excel file selected hasn't been found. Cash forecast is NOT updated in Access", vbCritical
                End If

                Rem if cash forecast tab + customer no. column exists:
                If Found Then
                    row = 0
                    col = 0
                    Found = False
                    For row = 1 To 10
                        For col = 1 To 100
                            If .Sheets(TabCashForecast).Cells(row, col) = UCase("cash forecast") Then
                                Found = True
                                Exit For
                            End If
                            If Found Then Exit For
                        Next col
                        If Found Then Exit For
                    Next row

                    If Found Then
                        CurrentDb.Execute "UPDATE Tbl_Customers SET Tbl_Customers.MonthlyTargetInMainCurrency = 0;"
                        Rem SELECT Tbl_CashCollected.CustomerID, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.Currency, Tbl_CashCollected.Amount, Tbl_Currencies.ExchangeRate, [Amount]*[ExchangeRate] AS AmountInEUR FROM Tbl_CashCollected LEFT JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_CashCollected.CustomerID)=1750) AND ((Tbl_CashCollected.[Payment Date]) Between #1/1/2013# And #1/10/2013#));
                        Rem CurrentDb.Execute "UPDATE Tbl_Customers SET Tbl_Customers.Customer_code = "xxxxxxxx", Tbl_Customers.MonthlyTargetInMainCurrency = 123;"

                        row = 5


                        Set rs1 = CurrentDb.OpenRecordset("SELECT Tbl_Customers.Customer_code, Tbl_Customers.MonthlyTargetInMainCurrency FROM Tbl_Customers;")
                        rs1.MoveFirst




                        While .Sheets(TabCashForecast).Cells(row, ColCustID) <> ""
                            Set rst = CurrentDb.OpenRecordset("SELECT Sum([ExchangeRate]*[Original amount]) AS AmountInEUR FROM Tbl_CashCollected INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_CashCollected.CustomerID)=" & .Sheets(TabCashForecast).Cells(row, ColCustID) & ") AND ((Tbl_CashCollected.[Payment Date]) Between #" & Format(Text0.value, "mm/dd/yy") & "# And #" & Format(Text2.value, "mm/dd/yy") & "#));")
                            rs1.MoveFirst
                            rs1.FindFirst "Customer_code=" & Trim(.Sheets(TabCashForecast).Cells(row, ColCustID))
                            If rs1.NoMatch = False Then
                                    rs1.Edit
                                    rs1.Fields("MonthlyTargetInMainCurrency") = Round((.Sheets(TabCashForecast).Cells(row, col)) + Nz(rst.Fields("AmountInEUR"), 0), 0)
                                    rs1.Update
                            End If
                            row = row + 1
                        Rem    Label9.Caption = row
                            DoEvents
                        Wend
                    Else
                        MsgBox "Column Cash forecast in the Excel file selected hasn't been found. Cash forecast is NOT updated in Access", vbCritical

                    End If
                End If
            End With
            ExcDoc.Close
            Set ExcDoc = Nothing
            Set ExcApp = Nothing
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_Load()
    Text0.value = DateAdd("d", 1, DMax("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd <#" & Format(Date, "mm/dd/yy") & "#"))
    Rem Text2.value = DMin("[MonthEnd]", "[Tbl_MonthEnd]", "MonthEnd >#" & Format(DateAdd("d", 1, CDate(Text0.value)), "mm/dd/yy") & "#")
End Sub


Private Sub Text0_Click()
    Command4.enabled = (Not IsNull(Text0.value)) And (Not IsNull(Text2.value))
End Sub
Private Sub Text2_Change()
    Call Text0_Click
End Sub

Private Sub Text2_Click()
    Call Text0_Click
End Sub
