Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "SELECT Tbl_Invoices.Currency, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date FROM Tbl_Invoices GROUP BY Tbl_Invoices.Currency, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date HAVING (((Tbl_Invoices.Update_date)=#" & Format(FindLastDate(), "mm/dd/yyyy") & "#));"
End Sub
