Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
Dim rst2, Rest3 As Recordset
Dim SQL As String
Rem If OpenArgs = 0 Then
Rem     Me.Graph1.RowSource = "SELECT Left([Name],25) AS CustomerName, (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE ( ((Tbl_Invoices.Update_date)=#" & format(Date, "mm/dd/yy") & "#) AND ((Tbl_queries.InvoiceToBePaid)=False)) GROUP BY Left([Name],25) ORDER BY (Sum([amount]*[exchangerate]))/1000 DESC;"
Rem Else
Rem    Me.Graph1.RowSource = "SELECT Left([Name],25) AS CustomerName, (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Invoices.Update_date) = #" & format(Date, "mm/dd/yy") & "#) And ((Tbl_queries.InvoiceToBePaid) = False) And ((Tbl_Customers.Credit_controller) = " & OpenArgs & ")) GROUP BY Left([Name],25) ORDER BY (Sum([amount]*[exchangerate]))/1000 DESC;"
Rem End If
Rem Set Rst = CurrentDb.OpenRecordset(Me.Graph1.RowSource)
Rem Me.Detail.Height = (Rst.RecordCount * 5)
Rem Me.Graph1.Height = Me.Detail.Height - 10
Dim MainCurrency As String
    MainCurrency = DLookup("MainCurrency", "TblGeneral")
    Me.Graph1.ChartTitle.Text = "All disputing customers " & Chr(13) & "(in thousands " & MainCurrency & ")"

    SQL = OpenArgs
    Set rst3 = CurrentDb.OpenRecordset(SQL)
    rst3.MoveFirst

    CurrentDb.Execute ("Delete * FROM Tbl_AllComplainingCustomers;")
    Set rst2 = New ADODB.Recordset
    rst2.ActiveConnection = CurrentProject.Connection
    rst2.Open "Tbl_AllComplainingCustomers", , adOpenKeyset, adLockOptimistic, adCmdTable

    While Not rst3.EOF
        rst2.AddNew
        rst2.Fields("CustomerName") = rst3.Fields("CustomerName")
        Rem "SELECT (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Customers.Customer_code)="numero cliente") AND ((Tbl_Customers.Credit_controller)=1) AND ((Tbl_Invoices.Update_date)=Date()) AND ((Tbl_queries.InvoiceToBePaid)=False) AND ((Tbl_Invoices.Overdue_Date)="scadenze"));
        Set rst = CurrentDb.OpenRecordset("SELECT (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Customers.Customer_code)=" & rst3.Fields("Customer_code") & ") AND  ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_queries.InvoiceToBePaid)=False) AND ((Tbl_Invoices.Overdue_Date)>=#" & Format(DateAdd("d", -30, Date), "mm/dd/yy") & "#));")
        rst2.Fields("Within 30 Days") = Nz(rst.Fields("AmountInEUR"), 0)

        Set rst = CurrentDb.OpenRecordset("SELECT (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Customers.Customer_code)=" & rst3.Fields("Customer_code") & ") AND  ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_queries.InvoiceToBePaid)=False) AND ((Tbl_Invoices.Overdue_Date) between #" & Format(DateAdd("d", -31, Date), "mm/dd/yy") & "# AND #" & Format(DateAdd("d", -60, Date), "mm/dd/yy") & "#));")
        rst2.Fields("Between 31 and 60 Days") = Nz(rst.Fields("AmountInEUR"), 0)

        Set rst = CurrentDb.OpenRecordset("SELECT (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Customers.Customer_code)=" & rst3.Fields("Customer_code") & ") AND  ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_queries.InvoiceToBePaid)=False) AND ((Tbl_Invoices.Overdue_Date) between #" & Format(DateAdd("d", -61, Date), "mm/dd/yy") & "# AND #" & Format(DateAdd("d", -90, Date), "mm/dd/yy") & "#));")
        rst2.Fields("Between 61 and 90 Days") = Nz(rst.Fields("AmountInEUR"), 0)

        Set rst = CurrentDb.OpenRecordset("SELECT (Sum([amount]*[exchangerate]))/1000 AS AmountInEUR FROM (Tbl_Customers INNER JOIN (Tbl_queries INNER JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query) ON Tbl_Customers.Customer_code = Tbl_Invoices.Customer_ID) INNER JOIN Tbl_Currencies ON Tbl_Invoices.Currency = Tbl_Currencies.CurrencyID WHERE (((Tbl_Customers.Customer_code)=" & rst3.Fields("Customer_code") & ") AND  ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_queries.InvoiceToBePaid)=False) AND ((Tbl_Invoices.Overdue_Date)<=#" & Format(DateAdd("d", -91, Date), "mm/dd/yy") & "#));")
        rst2.Fields("Over 90 days") = Nz(rst.Fields("AmountInEUR"), 0)

        rst3.MoveNext
    Wend
    rst2.Update
    rst2.Close
    rst2 = Null
    rst3 = Null
    Me.Requery
    Me.Graph1.Requery
    DoCmd.Maximize
End Sub
