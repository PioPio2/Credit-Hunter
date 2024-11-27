Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Delete(Cancel As Integer)
    On Error Resume Next
        Kill Me.Controls("AttachName")
    On Error GoTo 0
End Sub
