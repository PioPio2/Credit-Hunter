Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Descripition_DblClick(Cancel As Integer)
Dim rst As Recordset
On Error GoTo ErrHandling
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_DocumentsToBeErased", , adOpenKeyset, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("CustomerID") = Me.Form.Parent.Customer_code.value
        .Fields("DocumentType") = Me.Form.Recordset.Fields("ID")
        .Update
        Me.Form.Parent.[MskDocumentsToBeErasedSubform].Requery
        Me.Form.Parent.[MskDocumentsToBeErasedSubform].Form.Recordset.MoveFirst
        .Close
    End With
    Set rst = Nothing
    Exit Sub
ErrHandling:
If err.number = -2147217887 Then
    a = MsgBox("The document is already selected to be deleted.", vbCritical, "Error")
Else
    a = MsgBox(err.number & " " & err.Description, vbCritical)
End If
err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub
