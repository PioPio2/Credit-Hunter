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
    DoCmd.RunSQL "Delete Tbl_Link_Customer_Internal_Email_Address.CustomerID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID FROM Tbl_Link_Customer_Internal_Email_Address WHERE (((Tbl_Link_Customer_Internal_Email_Address.CustomerID)=" & Me.Form.Parent.Customer_code.value & ") AND ((Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID)=" & Me.Form.Recordset.Fields("InternalEmailAddressID") & "));"
    Me.Requery
End Sub
Public Function GetInternalEmailCC(TemplateNumber As Integer) As String
Dim rst As Variant
    If Me.Recordset.RecordCount > 0 Then
        Set rst = New ADODB.Recordset
        Me.Recordset.MoveFirst
        GetInternalEmailCC = ""
        While Not Me.Recordset.EOF
            Set rst = CurrentDb.OpenRecordset("SELECT TBL_LinkTemplateEmailAddress.IDTemplate, TBL_LinkTemplateEmailAddress.IDDepartment FROM TBL_LinkTemplateEmailAddress where (TBL_LinkTemplateEmailAddress.IDTemplate=" & TemplateNumber & ") and (TBL_LinkTemplateEmailAddress.IDDepartment =" & Me.Recordset.Fields("Department") & ");")
            If rst.RecordCount > 0 Then
                GetInternalEmailCC = GetInternalEmailCC & Me.Recordset.Fields("EmailAddress") & ","
            End If
            Me.Recordset.MoveNext
        Wend
        If GetInternalEmailCC <> "" Then
            GetInternalEmailCC = Left(GetInternalEmailCC, Len(GetInternalEmailCC) - 1)
        End If
        Set rst = Nothing
    End If
End Function
