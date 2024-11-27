Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private mPB_O2 As clsPROGRESS
Dim CountFile, LFile As Long
Private Sub InsertNewCreditLimit(I As Integer, nfile As Variant)
Dim Country, Inputdata, InputData2, CustCode, AreaName As String
Dim rst As Recordset
Dim RstCustomers As Recordset
Dim RstAreas As Recordset
Dim RstCountries As Recordset
Dim AreaCode As Long


If nfile <> "" Then
    DoCmd.OpenQuery ("QueryToReleaseOrder")
    DoCmd.OpenQuery ("QueryToSendReleaseRequest")

    Set rst = New ADODB.Recordset
    rst.ActiveConnection = CurrentProject.Connection
    rst.Open "Tbl_CL", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set RstCustomers = New ADODB.Recordset
    RstCustomers.ActiveConnection = CurrentProject.Connection
    RstCustomers.Open "Tbl_Customers", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set RstAreas = New ADODB.Recordset
    RstAreas.ActiveConnection = CurrentProject.Connection
    RstAreas.Open "Tbl_Areas", , adOpenKeyset, adLockOptimistic, adCmdTable

    Open nfile For Input As #1    ' Apre il file per l'input.
    Open nfile For Input As #2    ' Apre copia del file per l'input.
    Set mPB_O2 = New clsPROGRESS
    mPB_O2.CreaPB Me.shPB_O2, Orizzontale
    mPB_O2.DefinisciPB
    mPB_O2.Progressione (0)
    Me.Etichetta99.Visible = True
    Line Input #2, InputData2
    While Not EOF(1)
        Line Input #1, Inputdata
        If Not EOF(2) Then
            Line Input #2, InputData2
        End If
        CountFile = CountFile + Len(Inputdata) + 2
        If IsNumeric(Trim(Mid$(Inputdata, 45, 5))) Then
            With rst
                AreaName = Trim(Left(Inputdata, 12))
                If Not IsNumeric(Trim(Mid$(InputData2, 45, 5))) And (Asc(Left(InputData2, 1)) > 12) Then
                    AreaName = Trim(AreaName & " " & Trim(Left(InputData2, 12)))
                End If
                RstAreas.MoveFirst
                If Not RstAreas.EOF Then
                    RstAreas.MoveFirst
                    RstAreas.Find ("Area='" & AreaName & "'")
                End If
                If RstAreas.EOF Then
                    RstAreas.AddNew
                    RstAreas.Fields("Area") = AreaName
                    RstAreas.Update
                End If

                AreaCode = RstAreas.Fields("ID")

                If .RecordCount > 0 Then
                    .MoveFirst
                    '.Find ("Customer_code='" & Trim(Mid$(Inputdata, 45, 5)) & "'")
                    .Find ("Customer_code='" & Trim(Mid$(Inputdata, 45, 6)) & "'")
                End If
                If .EOF Then
                    .AddNew
                    '.Fields("Customer_code") = Trim(Mid$(Inputdata, 45, 5))
                    .Fields("Customer_code") = Trim(Mid$(Inputdata, 45, 6))
                End If
                'CustCode = Trim(Mid$(Inputdata, 45, 5))
                CustCode = Trim(Mid$(Inputdata, 45, 6))
                RstCustomers.MoveFirst
                RstCustomers.Find ("Customer_code='" & CustCode & "'")
                If Not RstCustomers.EOF Then
                    RstCustomers.Fields("Area") = AreaCode
                    RstCustomers.Update
                End If

                If I = 1 Then
                    .Fields("Currency") = Trim(Mid$(Inputdata, 51, 3))
                    If Trim(Mid$(Inputdata, 54, 17)) = "UNLIMITED" Then
                        .Fields("CreditLimit") = 99999999
                    Else
                        .Fields("CreditLimit") = Trim(Mid$(Inputdata, 54, 17))
                    End If
                    .Fields("OpenARBalance") = Trim(Mid$(Inputdata, 106, 14))
                    .Fields("AwaitingInvoicing") = Mid$(Inputdata, 124, 13)
                    .Fields("AmtScheduledTom") = Mid$(Inputdata, 150, 15)
                Else
                    .Fields("AmtScheduled8Dyas") = Mid$(Inputdata, 150, 15)
                End If
                mPB_O2.Progressione (CInt(100 * CountFile / LFile))
                Me.Etichetta99.Caption = (CStr(Int(100 * CountFile / LFile))) & "%"
                .Update
                Me.Repaint
                DoEvents
            End With
        End If
    Wend
    rst.Close
    RstCustomers.Close
    RstAreas.Close
    Set rst = Nothing
    Set RstAreas = Nothing
    Set RstCustomers = Nothing
    Set mPB_O2 = Nothing
    Me.Etichetta99.Visible = False
    Close #1
    Close #2



    Rem update import datetime
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        If I = 1 Then
            .Fields("Update_CL+1") = Now()
        Else
            .Fields("Update_CL+8") = Now()
        End If
        .Update
        .Close
    End With
    Set rst = Nothing
    MsgBox "Import credit limits completed", vbExclamation
    Me.Repaint
End If
End Sub

Private Sub Form_Timer()
Dim rst As Recordset
Dim Inputdata, StrInputFileName, StrFilter As String
    StrFilter = ahtAddFilterItem(StrFilter, "Text files(*.txt)", "*.txt")
    StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
        DialogTitle:="Please select the credit limit file to upload...", _
        Flags:=ahtOFN_HIDEREADONLY)
    Me.TimerInterval = 0
    If StrInputFileName <> "" Then
        Me.Etichetta99.Visible = True
        Me.Repaint
        CountFile = 0
        LFile = FileLen(StrInputFileName)
        Open StrInputFileName For Input As #1    ' Apre il file per l'input.
        While InStr(1, Inputdata, "Schedule Date Threshold") = 0
            Line Input #1, Inputdata
            CountFile = CountFile + Len(Inputdata) + 2
        Wend
        Close #1
        If CDate(Mid(Inputdata, 63, 10)) <> DateAdd("d", Date, 5) Then
            MsgBox "Please check the report you are uploading because the horizon date in it is different than what expected. However, I will proceed with the uploading now.", vbExclamation
        End If

        Rem Insert Horizon Date Limit
        Set rst = New Recordset
        With rst
            .ActiveConnection = CurrentProject.Connection
            .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
            .Fields("CLHorizonDateLimit") = CDate(Mid(Inputdata, 63, 10))
            .Update
            .Close
        End With
        Set rst = Nothing

        Call InsertNewCreditLimit(1, StrInputFileName)
        Me.Refresh
        Me.Requery

        Rem #########  Updates Historical Credit Limits #######
        Call UpdateHistoricalCL
        Label0.Visible = True
        Me.Repaint
        DoEvents
        Rem ######### SEND CL AUTOMATICALLY #########
        If Not (IsNull(DLookup("ToBeSentCLto", "TblGeneral"))) Then
            Call CLLimitReport
            Call SendEmails(DLookup("ToBeSentCLto", "TblGeneral"), "", "Updated Credit Limits", "", GetPathExcelDirectory() & "Updated CL Report.xls")
        End If
    End If
    DoCmd.Close acForm, "MskImportDailyCLs", acSaveNo
End Sub
