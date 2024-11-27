Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    Dim rst, rst2, bb As Recordset
    Dim I, a As Long
    I = 0
    a = 0
    Set rst2 = CurrentDb.OpenRecordset("SELECT Tbl_Invoices_History.* FROM Tbl_Invoices_History;")
    rst2.MoveLast
    rst2.MoveFirst
    Set rst = CurrentDb.OpenRecordset("SELECT DISTINCTROW Sheet1.Customer_ID, Sheet1.Document_Number, Tbl_Invoices.Amount, Max(Tbl_Invoices.Update_date) AS MaxOfUpdate_date FROM Sheet1 LEFT JOIN Tbl_Invoices ON (Sheet1.Document_Number = Tbl_Invoices.Document_Number) AND (Sheet1.Customer_ID = Tbl_Invoices.Customer_ID) GROUP BY Sheet1.Customer_ID, Sheet1.Document_Number, Tbl_Invoices.Amount;")
    With rst
        .MoveLast
        .MoveFirst
        While .EOF = False
            rst2.MoveFirst
            rst2.FindFirst "Customer_id=" & rst.Fields("Customer_ID") & " and Document_Number='" & rst.Fields("Document_Number") & "'"
            If rst2.NoMatch = False Then
                If (IsNull(rst2.Fields("PaymentDate"))) Then
                    a = a + 1
                    rst2.Edit
                    rst2.Fields("PaymentDate") = rst.Fields("MaxOfUpdate_date")
                    rst2.Update
                End If
            End If
            .MoveNext
            I = I + 1
            Label1.Caption = I & " of " & rst.RecordCount & " and updated records " & a
            DoEvents
        Wend
    End With
    rst.Close
End Sub
