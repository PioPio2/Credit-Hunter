Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me.ID = DMax("[ID]", "[Tbl_EmailAddresses]") * 2
End Sub
