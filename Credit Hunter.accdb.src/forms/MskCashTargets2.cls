Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
Dim Schedule As Byte
Dim rst As Recordset
    Schedule = 0
    If Check42.value <> 0 Then
        Schedule = Schedule + 2
    End If
    If Check44.value <> 0 Then
        Schedule = Schedule + 4
    End If
    If Check46.value <> 0 Then
        Schedule = Schedule + 8
    End If
    If Check48.value <> 0 Then
        Schedule = Schedule + 16
    End If
    If Check50.value <> 0 Then
        Schedule = Schedule + 32
    End If

    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "tblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
        .Fields.item("CashTargetReportSchedule") = Schedule
        .Update
Rem        .Save
        .Close
    End With
    Set rst = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If

End Sub

Private Sub Form_Load()
Dim Schedule As Byte
    Schedule = Nz(DLookup("CashTargetReportSchedule", "tblGeneral"), 0)
    If Schedule <> 0 Then
        Check42.value = Schedule And 2
        Check44.value = Schedule And 4
        Check46.value = Schedule And 8
        Check48.value = Schedule And 16
        Check50.value = Schedule And 32
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.MskSubTarget.Visible = False
End Sub

Private Sub Text1_Change()
    Label4.Visible = True
    Text3.Visible = True
    Call FilterOnSubmask
End Sub

Private Sub FilterOnSubmask()
Dim Filter As String
Me.Application.Echo False
Me.MskSubTarget.Form.FilterOn = False
Filter = ""
If Label2.Visible = True Then
    Filter = "FiscalYear=" & CInt(Me.Text1.value)
End If
If (Text3.Visible = True) And Not (IsNull(Text3.value)) Then
    Filter = Filter & " AND FiscalQuarter=" & CInt(Me.Text3.value)
End If
If (Text5.Visible = True) And Not (IsNull(Text5.value)) Then
    Filter = Filter & " AND FiscalMonth=" & CInt(Me.Text5.value)
    Me.MskSubTarget.Visible = True
End If

If Filter <> "" Then
    Me.MskSubTarget.Form.Filter = Filter
    Me.MskSubTarget.Form.FilterOn = True

Else
    Me.MskSubTarget.Visible = False
End If
Me.Application.Echo True
End Sub

Private Sub Text3_Change()
Dim I As Integer
    Label6.Visible = True
    Text5.Visible = True
    If Text3.value = 4 Then
        Text5.RowSource = "1,2,3"
    Else
        Text5.RowSource = ""
        For I = 1 To 3
            Text5.RowSource = Text5.RowSource & CStr(Text3.value * 3 + I) & ","
        Next I
    End If
    Call FilterOnSubmask
End Sub

Private Sub Text5_Change()
    Call FilterOnSubmask
End Sub
