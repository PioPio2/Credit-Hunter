Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
    Call PutOrder
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub

Private Sub Form_Load()
    Call PutOrder
End Sub
Private Sub Form_Open(Cancel As Integer)
    Call PutOrder
End Sub
Sub PutOrder()
End Sub

Private Sub Note_DblClick(Cancel As Integer)

DoCmd.OpenForm "MskMagnifier", acNormal, , , , acDialog, Note.value
End Sub
