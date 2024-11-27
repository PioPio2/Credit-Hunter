Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Timer()
Dim Str As Variant
Dim S, Str1, StrFilter As String
Dim NumFiles, I, ArrayDimension As Integer
Dim AllTXT2 As Boolean
Dim TextFiles() As String
    StrFilter = ahtAddFilterItem(StrFilter, "Excel files(*.xls, *.xlsx) or Text file (*.txt)", "*.xls;*.xlsx;*.txt")
    Str1 = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select Credit check failure file to upload...", _
                Flags:=ahtOFN_HIDEREADONLY Or ahtOFN_ALLOWMULTISELECT Or ahtOFN_EXPLORER)
    NumFiles = 0
    If IsArray(Str1) Then
        NumFiles = UBound(Str1) + 1
    ElseIf Str1 <> "" Then
        NumFiles = 1
    End If

    If NumFiles = 1 Then
        ReDim Str(0)
        Str(0) = Str1
    End If
    If NumFiles > 0 Then
        DoCmd.OpenQuery ("QueryEraseTblCreditCheckFailures")
        DoCmd.OpenQuery ("QueryToReleaseOrder")
        DoCmd.OpenQuery ("QueryToSendReleaseRequest")

        For I = 0 To NumFiles - 1 ' insert .txt files selected in a dedicated array
            If NumFiles = 1 Then
                S = Str1
            Else
                S = Str1(I)
            End If
            If Mid(S, InStrRev(S, ".") + 1) = "txt" Then
                On Error Resume Next
                ArrayDimension = UBound(TextFiles)
                If err.number = 0 Then
                    ReDim Preserve TextFiles(UBound(TextFiles) + 1)
                    ArrayDimension = UBound(TextFiles)
                Else
                    ReDim TextFiles(1)
                    ArrayDimension = UBound(TextFiles)
                End If
                TextFiles(UBound(TextFiles)) = S
            Else
                If CheckIfInShippingReportLayoutIsCorrect(S) = True Then
                    Call InsertCustomersFailFromExcelFile(S)
                    Label1.Caption = "Uploading customers failing the credit check..."
                    DoEvents
                ElseIf CheckIfOnHoldReportLayoutIsCorrect(S) = True Then
                    Label1.Caption = "Uploading customers on manual credit hold..."
                    DoEvents
                    InsertOnHoldCustomersFile (S)
                Else
                    MsgBox "The file " & S & " layout is neither an InShipping report or a On Hold report. Please double check it and do the upload again. This upload stops here.", vbCritical
                    DoCmd.Close acForm, "MskImportLeilaniFile2", acSaveNo
                    Exit Sub
                End If
            End If
        Next I
        If ArrayDimension > 0 Then
            S = MergeTXTFiles(TextFiles, 1) 'merge more than one hold report
            Label1.Caption = "Uploading customers failing the credit check..."
            DoEvents
            Call InsertCustomersFailFromTextFile(S)
        End If
        Label1.Caption = "Deleting duplicate lines..."
        DoEvents
        Rem CurrentDb.Execute "DELETE Tbl_credit_check_failures.ID FROM Tbl_credit_check_failures WHERE (((Tbl_credit_check_failures.ID) In (SELECT First(Tbl_credit_check_failures.ID) AS FirstOfID FROM Tbl_credit_check_failures GROUP BY Tbl_credit_check_failures.[Order Number], Tbl_credit_check_failures.[Order Line Number], Tbl_credit_check_failures.[Customer Name] HAVING (((First(Tbl_credit_check_failures.ID)) Not In (SELECT First(Tbl_credit_check_failures.ID) AS FirstOfID FROM Tbl_credit_check_failures GROUP BY Tbl_credit_check_failures.[Order Number], Tbl_credit_check_failures.[Order Line Number], Tbl_credit_check_failures.[Hold Name] HAVING (((Tbl_credit_check_failures.[Hold Name])='Credit Check Failure'));)) AND ((Count([Order Number] & [Order Line Number]))>1)))));"
        CurrentDb.Execute "DELETE Tbl_credit_check_failures.ID, [Order Line Number] & [Order Number] AS a, Tbl_credit_check_failures.[Hold Name] FROM Tbl_credit_check_failures WHERE ((([Order Line Number] & [Order Number]) In (SELECT [Order Line Number] & [Order Number] AS a FROM Tbl_credit_check_failures WHERE (((Tbl_credit_check_failures.[Hold Type])='LOGI Manual Credit Hold'));)) AND ((Tbl_credit_check_failures.[Hold Name])='Credit Check Failure'));"

        MsgBox "Upload completed succesfully.", vbExclamation
    End If

    DoCmd.Close acForm, "MskImportLeilaniFile2", acSaveNo
End Sub
Private Function CheckIfOnHoldReportLayoutIsCorrect(ByVal StrInputFileName As String) As Boolean
Header = Array("Order&LineID", "Hold Name", "Days on Hold", "Date Hold Applied", "Hold Comments", "Sub-Region", "Country", "Customer Name", "Customer Number", "Account Specialist", "Logitech Item Number", "Marketing Name", "Customer Promotional Reference", "Price List Name", "Special Price", "Requested Quantity", "Currency Code", "Amount", "EUR Amount", "Order Number", "Order Line Number", "Order Date", "Requested Date", "Schedule Date", "Release Set", "Line Status", "Hold Criteria", "Tax Code", "EUR Amount for high level overview")
Dim I As Integer
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook

CheckIfOnHoldReportLayoutIsCorrect = False
If StrInputFileName <> "" Then
    Set ExcApp = CreateObject("Excel.Application")
    Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
    ExcApp.Visible = False
    If ExcelSheetExists("Hold Report", ExcDoc) Then
        CheckIfOnHoldReportLayoutIsCorrect = True
        With ExcDoc.Sheets("Hold Report")
            For I = 0 To UBound(Header)
                If Header(I) <> .Cells(1, I + 1) Then
                    CheckIfOnHoldReportLayoutIsCorrect = False
                    Exit For
                End If
            Next I
        End With
    End If
End If
ExcDoc.Close SaveChanges:=False
Set ExcDoc = Nothing
Set ExcApp = Nothing

End Function

Private Function CheckIfInShippingReportLayoutIsCorrect(ByVal StrInputFileName As String) As Boolean
Header = Array("Hold Type", "Hold Name", "Date Hold Applied", "Hold Until Date", "Hold Comments", "Sub-Region", "Country", "Customer Name", "Customer Number", "Account Specialist", "Logitech Item Number", "Item Description", "List Price", "Requested Quantity", "Currency Code", "Amount", "Order Number", "Order Line Number", "Order Date", "Requested Date", "Schedule Date", "Active Hold", "Open Line", "Line Status", "Hold Criteria", "Tax Code")
Dim I As Integer
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook

CheckIfInShippingReportLayoutIsCorrect = False
If StrInputFileName <> "" Then
    Set ExcApp = CreateObject("Excel.Application")
    Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
    If ExcelSheetExists("Credit Check Failures", ExcDoc) Then
        CheckIfInShippingReportLayoutIsCorrect = True
        ExcApp.Visible = False
        With ExcDoc.Sheets("Credit Check Failures")
            For I = 0 To UBound(Header)
                If Header(I) <> .Cells(1, I + 1) Then
                    CheckIfInShippingReportLayoutIsCorrect = False
                    Exit For
                End If
            Next I
        End With
    End If
End If
ExcDoc.Close SaveChanges:=False
Set ExcDoc = Nothing
Set ExcApp = Nothing
End Function
Private Sub InsertOnHoldCustomersFile(ByVal StrInputFileName As String)
Dim rst As Recordset
Dim Riga, I As Integer
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim NRows As Integer

    If StrInputFileName <> "" Then
        Set mPB_O2 = New clsPROGRESS
        mPB_O2.CreaPB Me.shPB_O2, Orizzontale
        mPB_O2.DefinisciPB
        shPB_O2.Visible = True
        NRows = NumMaxRows(CStr(StrInputFileName), "Hold Report") - 1
        Set rst = New ADODB.Recordset
        rst.ActiveConnection = CurrentProject.Connection
        rst.Open "Tbl_credit_check_failures", , adOpenKeyset, adLockOptimistic, adCmdTable
        Set ExcApp = CreateObject("Excel.Application")
        Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
        ExcApp.Visible = False
        Riga = 2

        With ExcDoc.Sheets("Hold Report")
            shPB_O2.Visible = True
            Etichetta99.Visible = True
            While .Cells(Riga, 1) <> ""
                If (.Cells(Riga, 2) = "LOGI Manual Credit Hold") And (.Cells(Riga, 25) = "Y") Then
                    rst.AddNew

                    rst.Fields(1) = .Cells(Riga, 2) 'Hold Name
                    rst.Fields(4) = .Cells(Riga, 4) 'Hold comments
                    rst.Fields(5) = .Cells(Riga, 5) 'Sub-Region
                    rst.Fields(6) = .Cells(Riga, 6) 'Country
                    rst.Fields(7) = .Cells(Riga, 7) 'Customer Name
                    rst.Fields(8) = .Cells(Riga, 8) 'Customer Number
                    rst.Fields(9) = .Cells(Riga, 9) 'Account Specialist
                    rst.Fields(10) = .Cells(Riga, 10) 'Logitech Item Number
                    rst.Fields(11) = .Cells(Riga, 11) 'Marketing Name/item description
                    rst.Fields(12) = .Cells(Riga, 13) 'Price List Name
                    rst.Fields(13) = .Cells(Riga, 15) 'Requested Quantity
                    rst.Fields(14) = .Cells(Riga, 16) 'Currency Code
                    rst.Fields(15) = .Cells(Riga, 17) 'Amount
                    rst.Fields(16) = .Cells(Riga, 19) 'Order Number
                    rst.Fields(17) = .Cells(Riga, 20) 'Order Line Number
                    rst.Fields(18) = .Cells(Riga, 21) 'Order Date
                    rst.Fields(19) = .Cells(Riga, 22) 'Requested Date
                    rst.Fields(20) = .Cells(Riga, 23) 'Schedule Date
                    rst.Fields(23) = .Cells(Riga, 25) 'Line Status
                    rst.Fields(24) = .Cells(Riga, 26) 'Hold Criteria
                    rst.Fields(25) = .Cells(Riga, 27) 'Tax Code

                    rst.Update
                End If
                mPB_O2.Progressione (CInt(100 * Riga / NRows))
                Etichetta99.Caption = (CInt(100 * Riga / NRows)) & "%"
                Riga = Riga + 1
                Me.Repaint
                DoEvents
            Wend
        End With

        ExcDoc.Close SaveChanges:=False
        Set ExcDoc = Nothing
        Set ExcApp = Nothing
        rst.Close

        Rem update import datetime
        Set rst = New ADODB.Recordset
        With rst
            .ActiveConnection = CurrentProject.Connection
            .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
            .MoveFirst
            .Fields("Update_Customers_Failing") = Now()
            .Update
            .Close
        End With
        Set rst = Nothing

        Set mPB_O2 = Nothing
        Me.Refresh
        Me.Requery

    End If

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    Exit Sub
End Sub


Private Sub InsertCustomersFailFromExcelFile(ByVal StrInputFileName As String)
Dim rst As Recordset
Dim Riga, I As Integer
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim Header
Dim NRows As Integer

Header = Array("Hold Type", "Hold Name", "Date Hold Applied", "Hold Until Date", "Hold Comments", "Sub-Region", "Country", "Customer Name", "Customer Number", "Account Specialist", "Logitech Item Number", "Item Description", "List Price", "Requested Quantity", "Currency Code", "Amount", "Order Number", "Order Line Number", "Order Date", "Requested Date", "Schedule Date", "Active Hold", "Open Line", "Line Status", "Hold Criteria", "Tax Code")
    If StrInputFileName <> "" Then
        Set mPB_O2 = New clsPROGRESS
        mPB_O2.CreaPB Me.shPB_O2, Orizzontale
        mPB_O2.DefinisciPB
        shPB_O2.Visible = True
        NRows = NumMaxRows(CStr(StrInputFileName), "Credit Check Failures")
        Set rst = New ADODB.Recordset
        rst.ActiveConnection = CurrentProject.Connection
        rst.Open "Tbl_credit_check_failures", , adOpenKeyset, adLockOptimistic, adCmdTable
        Set ExcApp = CreateObject("Excel.Application")
        Set ExcDoc = ExcApp.Workbooks.Open(StrInputFileName)
        ExcApp.Visible = False
        Riga = 2
        With ExcDoc.Sheets("Credit Check Failures")
            shPB_O2.Visible = True
            Etichetta99.Visible = True
            While .Cells(Riga, 1) <> ""
                If .Cells(Riga, 24) = "Awaiting Shipping" Or .Cells(Riga, 24) = "Picked Partial" Then
                    rst.AddNew
                    For I = 0 To 24
                        rst.Fields(Header(I)) = .Cells(Riga, I + 1)
                    Next I
                    rst.Update
                End If
                mPB_O2.Progressione (CInt(100 * Riga / NRows))
                Etichetta99.Caption = (CInt(100 * Riga / NRows)) & "%"
                Riga = Riga + 1
                Me.Repaint
                DoEvents
            Wend
        End With

        ExcDoc.Close SaveChanges:=False
        Set ExcDoc = Nothing
        Set ExcApp = Nothing
        rst.Close

        Rem update import datetime
        Set rst = New ADODB.Recordset
        With rst
            .ActiveConnection = CurrentProject.Connection
            .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
            .MoveFirst
            .Fields("Update_Customers_Failing") = Now()
            .Update
            .Close
        End With
        Set rst = Nothing

        Set mPB_O2 = Nothing
        Me.Refresh
        Me.Requery

    End If
    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    Exit Sub
End Sub

Private Sub InsertCustomersFailFromTextFile(ByVal Str As String)
    Dim TextLine, DirSave, Inputdata As String
    Dim b, Count As Integer
    shPB_O2.Visible = False
    Label1.Visible = True
    Me.Etichetta99.Visible = False

    '------------------------------
    Label1.Caption = "Phase 1: Parsing file..."
    DoEvents
    Open Str For Input As #1    ' Apre il file per l'input.
    Line Input #1, TextLine
    If DLookup("CreditLimitCheckFailureTXTFileHeader", "TblGeneral") = TextLine Then
        Close #1
        Open Str For Input As #1    ' Apre il file per l'input.
        DirSave = "C:\Users\" & fOSUserName() & "\DOCUMENTS\CreditCheckFailureTXTFileAmended.txt"
        Open DirSave For Output As #2    ' Apre il file per l'input.
        Count = 1
        While Not EOF(1)
            Line Input #1, Inputdata
            If Count = 1 Then
                Count = 2
                Inputdata = Replace(Inputdata, ",", "¬")
            Else
                Inputdata = Replace(Inputdata, Chr(34) & "," & Chr(34), "¬")
                Inputdata = (Left(Inputdata, Len(Inputdata) - 1))
                Inputdata = (Right(Inputdata, Len(Inputdata) - 1))
            End If
            Print #2, Inputdata
        Wend
        Close #1
        Close #2
        '----------------------
        CurrentDb.Execute "DELETE CreditCheckFailureTXTFileAmended.* FROM CreditCheckFailureTXTFileAmended;" ' delete Tbl_LOGI_Hold_Report_New

        DoCmd.TransferText acImportDelim, "Import-CreditCheckFailureTXTFileAmended", "CreditCheckFailureTXTFileAmended", DirSave, True, ""  ' import Credit check failure lines from .txt file
        CurrentDb.Execute "DELETE CreditCheckFailureTXTFileAmended.* FROM CreditCheckFailureTXTFileAmended WHERE (((CreditCheckFailureTXTFileAmended.[Hold Name]) Not In (SELECT HoldTypesToBeConsideredForCreditCheckReleases FROM Tbl_HoldTypesToBeConsideredForCreditCheckReleases)));" 'delete lines not to consider in credit check release



        Label1.Caption = "Phase 2: Uploading data..."
        DoEvents

        'populates Tbl_credit_check_failures table with data just imported
        CurrentDb.Execute "INSERT INTO Tbl_credit_check_failures ( [Hold Type], [Hold Name], [Date Hold Applied], [Hold Until Date], [Hold Comments], [Sub-Region], Country, [Customer Name], [Customer Number], [Account Specialist], [Logitech Item Number], [Item Description], [List Price], [Requested Quantity], [Currency Code], Amount, [Order Number], [Order Line Number], [Order Date], [Requested Date], [Schedule Date], [Active Hold], [Open Line], [Line Status], [Hold Criteria], [Tax Code] ) " & _
        "SELECT CreditCheckFailureTXTFileAmended.[Hold Type], CreditCheckFailureTXTFileAmended.[Hold Name], CreditCheckFailureTXTFileAmended.[Date Hold Applied], CreditCheckFailureTXTFileAmended.[Hold Until Date], CreditCheckFailureTXTFileAmended.[Hold Comments], CreditCheckFailureTXTFileAmended.[Sub-Region], CreditCheckFailureTXTFileAmended.Country, CreditCheckFailureTXTFileAmended.[Customer Name], CreditCheckFailureTXTFileAmended.[Customer Number], CreditCheckFailureTXTFileAmended.[Account Specialist], CreditCheckFailureTXTFileAmended.[Logitech Item Number], CreditCheckFailureTXTFileAmended.[Item Description], CreditCheckFailureTXTFileAmended.[List Price], CreditCheckFailureTXTFileAmended.[Requested Quantity], CreditCheckFailureTXTFileAmended.[Currency Code], CreditCheckFailureTXTFileAmended.Amount, CreditCheckFailureTXTFileAmended.[Order Number], CreditCheckFailureTXTFileAmended.[Order Line Number], CreditCheckFailureTXTFileAmended.[Order Date], CreditCheckFailureTXTFileAmended.[Requested Date], " & _
        " CreditCheckFailureTXTFileAmended.[Schedule Date] , CreditCheckFailureTXTFileAmended.[Active Hold], CreditCheckFailureTXTFileAmended.[Open Line], CreditCheckFailureTXTFileAmended.[Line Status], CreditCheckFailureTXTFileAmended.[Hold Criteria], CreditCheckFailureTXTFileAmended.[Tax Code] FROM CreditCheckFailureTXTFileAmended;"

        Label1.Visible = False

        Rem b = MsgBox("Credit check failures uploading completed.", vbOKOnly)
    Else
        Close #1
        b = MsgBox("The layout of the file you have chosen seems not to be the correct one. Please check it and upload it again", vbOKOnly)
    End If

End Sub
