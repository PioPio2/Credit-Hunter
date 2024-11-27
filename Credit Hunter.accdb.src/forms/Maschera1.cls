Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Activate()
    DoCmd.Restore
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim b As Integer
    b = vbNo
    While (GetImportingProcess = True) And (b = vbNo)
        b = MsgBox("Upload ongoing. Click YES if you go ahead but some data could be incorrect, otherwise click NO to check again the status.", vbYesNo, "Warning")
    Wend
    Rem DoCmd.Maximize
End Sub
