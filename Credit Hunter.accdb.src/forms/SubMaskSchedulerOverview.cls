Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Channel_Click()
    ID_Click
End Sub

Private Sub Channel_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub

Private Sub ID_Click()
    DoCmd.GoToRecord acDataForm, "MskScheduler3", acGoTo, Me.CurrentRecord
End Sub

Private Sub ID_DblClick(Cancel As Integer)
    DoCmd.GoToRecord acDataForm, "MskScheduler3", acGoTo, Me.CurrentRecord
    Me.Parent.TabCtl48.value = 1
End Sub

Private Sub Name_Click()
    ID_Click
End Sub

Private Sub Name_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub

Private Sub Status_Click()
    ID_Click
End Sub

Private Sub Status_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub

Private Sub Text14_Click()
    ID_Click
End Sub

Private Sub Text14_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub

Private Sub Text16_Click()
    ID_Click
End Sub

Private Sub Text16_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub

Private Sub Text18_Click()
    ID_Click
End Sub

Private Sub Text18_DblClick(Cancel As Integer)
    ID_DblClick (Cancel)
End Sub
