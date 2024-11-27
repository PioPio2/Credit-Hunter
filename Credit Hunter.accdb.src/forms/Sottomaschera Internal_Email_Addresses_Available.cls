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

Private Sub Testo4_DblClick(Cancel As Integer)
Dim rst As Recordset
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Link_Customer_Internal_Email_Address", , adOpenKeyset, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("CustomerID") = Me.Form.Parent.Customer_code.value
        .Fields("InternalEmailAddressID") = Me.Form.Recordset.Fields("ID")
        .Update
        Me.Requery
        Me.Form.Parent.[Sottomaschera Tbl_Link_Customer_Internal_Email_Address].Requery
    End With
    Set rst = Nothing
End Sub
