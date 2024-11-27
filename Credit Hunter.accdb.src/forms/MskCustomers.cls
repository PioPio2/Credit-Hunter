Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Activate()
    DoCmd.Restore
End Sub

Private Sub Form_Current()
Rem    Me.[Sottomaschera Tbl_Link_Customer_Internal_Email_Address_Avaialbe].Form.RecordSource = "SELECT Tbl_EmailAddresses.EmailAddress, Tbl_EmailAddresses.Department, Tbl_EmailAddresses.ID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID FROM Tbl_EmailAddresses LEFT JOIN Tbl_Link_Customer_Internal_Email_Address ON Tbl_EmailAddresses.ID = Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID WHERE (((Tbl_Link_Customer_Internal_Email_Address.CustomerID) Is Null Or (Tbl_Link_Customer_Internal_Email_Address.CustomerID) <> " & Me.Customer_code.Value & ")) GROUP BY Tbl_EmailAddresses.EmailAddress, Tbl_EmailAddresses.Department, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID, Tbl_EmailAddresses.ID;"
    Me.[Sottomaschera Tbl_Link_Customer_Internal_Email_Address_Avaialbe].Form.RecordSource = "SELECT Tbl_EmailAddresses.EmailAddress, Tbl_EmailAddresses.Department, Tbl_EmailAddresses.ID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID, Tbl_DepartmentNames.DeparmentName FROM (Tbl_EmailAddresses LEFT JOIN Tbl_Link_Customer_Internal_Email_Address ON Tbl_EmailAddresses.ID = Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID) LEFT JOIN Tbl_DepartmentNames ON Tbl_EmailAddresses.Department = Tbl_DepartmentNames.ID WHERE (((Tbl_Link_Customer_Internal_Email_Address.CustomerID) Is Null Or (Tbl_Link_Customer_Internal_Email_Address.CustomerID) <> " & Me.Customer_code.value & ")) GROUP BY Tbl_EmailAddresses.EmailAddress, Tbl_EmailAddresses.Department, Tbl_EmailAddresses.ID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID, Tbl_DepartmentNames.DeparmentName ORDER BY Tbl_EmailAddresses.Department, Tbl_EmailAddresses.EmailAddress;"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not IsNull(OpenArgs) Then
        Me.FilterOn = False
        Me.Filter = "Customer_code=" & OpenArgs
        Me.FilterOn = True
    End If
End Sub
