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

Private Sub Form_Open(Cancel As Integer)
    Label1.Caption = "Cash Targets in EUR currency to be collected within: " & Format(DMin("[MonthEnd]", "[Tbl_MonthEnd]", "[MonthEnd] >=   Date()"), "dd-mmm-yyyy")
End Sub
