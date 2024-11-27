Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Activate()
    DoCmd.Restore
    Label5.Caption = "Email Main Recipients who will receive the report starting from " & _
    Text6.value & " days before the fiscal month end."
End Sub


Private Sub Text6_LostFocus()
    Label5.Caption = "Email Main Recipients who will receive the report starting from " & _
    Text6.value & " days before the fiscal month end."
End Sub
