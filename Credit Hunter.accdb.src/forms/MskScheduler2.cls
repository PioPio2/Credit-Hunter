Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Comando13_Click()
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
        Me.RecordSource = "SELECT * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY Tbl_Customers.Index DESC;"
    End If
End Sub

Private Sub Comando21_Click()
Dim a As String
Dim rs1 As Recordset
Dim RsBanks   As Recordset
Dim RS As DAO.Recordset
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

    a = ExcelStatement(rs1, RS, RsBanks)
    If Cornice34.value = 2 Then
        SendEmails Me.Recordset.Fields("email"), IIf(IsNull(Me.Recordset.Fields("ccEmail")), "", Me.Recordset.Fields("ccEmail")), "Updated statement " & Testo3.value & " as of " & Date, Me.Text46.value, a

        Testo10.value = "Sent updated statement"
        Call Comando13_Click
    End If
    Rem rs1.Close
    Set rs1 = Nothing
    Rem rs.Close
    Set RS = Nothing
End Sub

Private Sub Comando41_Click()
    Dim S As String
    S = "SELECT * FROM Tbl_Customers WHERE Tbl_Customers.Name LIKE '*" & Testo42.value & "*';"
    Me.RecordSource = S
    If Me.Recordset.RecordCount = 0 Then
        ' se proprio non si trova niente
        MsgBox "No customer found. " & Chr(13) & "Check and try again.", vbExclamation, "Attention"
        Me.RecordSource = "SELECT * FROM Tbl_Customers;"
    End If
End Sub


Private Sub Form_AfterUpdate()
    If QueryFileTochange = True Then
        Call UpdateQueryFile(Me.Recordset)
    End If
End Sub

Private Sub Form_Current()
Dim r, rr As Recordset
    Me.Sottomaschera_TblNotes.Form.PutOrder
    Me.Caption = "Last update: " & Me.Testo60.value
    If Hour(Me.Testo60.value - Now) >= 8 Then
        Me.Caption = "!!! ATTENTION !!! " & Me.Caption
    End If
    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date) <=#" & Format(Date, "mm/dd/yy") & "#)) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta63.Caption = Format("0")
    Else
        Etichetta63.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=Date() And (Tbl_Invoices.Overdue_Date)>=DATEADD('D',-30,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta65.Caption = Format("0")
    Else
        Etichetta65.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-31,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEADD('D',-60,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta67.Caption = Format("0")
    Else
        Etichetta67.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-61,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEadd('D',-90,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta69.Caption = Format("0")
    Else
        Etichetta69.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-91,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEadd('D',-180,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta71.Caption = Format("0")
    Else
        Etichetta71.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-181,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='EUR'));")
    If r.RecordCount = 0 Then
        Etichetta73.Caption = Format("0")
    Else
        Etichetta73.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If







    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date) <=#" & Format(Date, "mm/dd/yy") & "#)) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta75.Caption = Format("0")
    Else
        Etichetta75.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=Date() And (Tbl_Invoices.Overdue_Date)>=DATEADD('D',-30,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta77.Caption = Format("0")
    Else
        Etichetta77.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-31,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEADD('D',-60,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta79.Caption = Format("0")
    Else
        Etichetta79.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-61,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEadd('D',-90,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta81.Caption = Format("0")
    Else
        Etichetta81.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-91,Date()) And (Tbl_Invoices.Overdue_Date)>=DATEadd('D',-180,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta83.Caption = Format("0")
    Else
        Etichetta83.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

    Set r = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date, Tbl_Invoices.Currency,Sum(Tbl_Invoices.Amount) AS SommaDiAmount FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID) =" & Testo1.value & ") And ((Tbl_Invoices.Overdue_Date)<=DATEADD('D',-181,Date()))) GROUP BY Tbl_Invoices.Currency,Tbl_Invoices.Update_date HAVING ((((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#)  and (Tbl_Invoices.Currency)='USD'));")
    If r.RecordCount = 0 Then
        Etichetta85.Caption = Format("0")
    Else
        Etichetta85.Caption = Format(r.Fields("SommaDiAmount"), "##,##0.00")
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "SELECT * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now()))ORDER BY Tbl_Customers.Index DESC;"
    DoCmd.Maximize
    Me.Sottomaschera_TblNotes.Form.AllowAdditions = False
    QueryFileTochange = False
End Sub

Private Sub Testo14_AfterUpdate()
    Me.RecordSource = "SELECT * FROM Tbl_Customers WHERE (((Tbl_Customers.Credit_controller)=" & GetNumCreditController(fOSUserName()) & ") AND ((Tbl_Customers.NextAppointment)<=Now()))ORDER BY Tbl_Customers.Index DESC;"
End Sub
