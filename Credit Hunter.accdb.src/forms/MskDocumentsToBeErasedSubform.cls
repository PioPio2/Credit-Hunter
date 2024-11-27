Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Text4_DblClick(Cancel As Integer)
    CurrentDb.Execute "DELETE Tbl_DocumentsToBeErased.*, Tbl_DocumentsToBeErased.CustomerID, Tbl_DocumentsToBeErased.DocumentType FROM Tbl_DocumentsToBeErased WHERE (((Tbl_DocumentsToBeErased.CustomerID)=" & Me.Form.Parent.Customer_code.value & ") AND ((Tbl_DocumentsToBeErased.DocumentType)=" & Me.Form.Recordset.Fields("DocumentType") & "));"
    Me.Requery
    Me.Form.Recordset.MoveFirst
End Sub
