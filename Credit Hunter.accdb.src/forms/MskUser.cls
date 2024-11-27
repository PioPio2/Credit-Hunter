Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Combo26_AfterUpdate()
    Call ChangePeriod(Text18, Combo26, Text33)
End Sub

Private Sub Combo27_AfterUpdate()
    Call ChangePeriod(Text19, Combo27, Text34)
End Sub


Private Sub Text16_AfterUpdate()
    Call ChangePeriod(Text16, Combo24, Text32)
End Sub

Private Sub ChangePeriod(number, period, field As Variant)

    field.value = Right$("00" & number.value, 2) & period.Column(1) & CLng(DateAdd(period.Column(1), number.value, Date))
End Sub

Private Sub Text18_AfterUpdate()
    Call ChangePeriod(Text18, Combo26, Text33)
End Sub

Private Sub Text19_AfterUpdate()
    Call ChangePeriod(Text19, Combo27, Text34)
End Sub

Private Sub Text51_BeforeUpdate(Cancel As Integer)
    If Text51.value <> Text53.value Then
        Cancel = True
        MsgBox "The email password has to be the same as the one in the box 'Re-type email password. Please check.", vbOKOnly + vbCritical, "Error: "
    End If

End Sub

Private Sub Text53_BeforeUpdate(Cancel As Integer)
    If Text51.value <> Text53.value Then
        Cancel = True
        MsgBox "The email password has to be the same as the one in the box 'Re-type email password. Please check.", vbOKOnly + vbCritical, "Error: "
    End If
End Sub
