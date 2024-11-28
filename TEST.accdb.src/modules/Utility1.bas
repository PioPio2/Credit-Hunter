Attribute VB_Name = "Utility1"
Option Compare Database


Option Explicit
Public SpeseLegaliRecuperate, ImportoPagato, CodiceAvvocato, NomeAvvocato, ParcellaAvvocato, DataInizioProceduraConcorsuale As String
Public DataPagamentoPraticaLegale, DataParcellaAvvocato As Date
Public ImportoParcellaAvvocato As Currency
Public VisualizzaintestazionePaginaRptRealOverdue As Boolean
Public QueryFileTochange As Boolean
Public ChargebackFileTochange As Boolean
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function apiFindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal strClass As String, _
    ByVal lpWindow As String) As Long

Private Declare PtrSafe Function apiSendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal _
    wParam As Long, lParam As Long) As Long

Private Declare PtrSafe Function apiSetForegroundWindow Lib "user32" Alias _
    "SetForegroundWindow" (ByVal hwnd As Long) As Long

Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias _
    "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare PtrSafe Function apiIsIconic Lib "user32" Alias _
    "IsIconic" (ByVal hwnd As Long) As Long

Dim qdef As DAO.QueryDef
Dim RS As DAO.Recordset
Dim CustomerLastDate As Date
Public NextMonthEnd As Date
Function NumMaxRows(Path As String, SheetName As String, Optional Start) As Integer
    Dim MyXl As Excel.Application
    Dim Min, Max As Long
    Dim SheetNumber, I As Integer
    Min = 1
    Max = 65000
    Set MyXl = CreateObject("excel.application")
Rem    path = "c:\a\a.xls"
    MyXl.Workbooks.Open Path
    MyXl.Visible = False
    SheetNumber = 0
    If SheetName = "" Then
        SheetNumber = 1
    Else
        For I = 1 To MyXl.ActiveWorkbook.Sheets.Count
            If MyXl.ActiveWorkbook.Sheets(I).Name = SheetName Then
                SheetNumber = I
            End If
        Next I
    End If
    If SheetNumber = 0 Then
        NumMaxRows = 0
        Exit Function
    End If

    MyXl.ActiveWorkbook.Sheets(SheetNumber).Select

    If IsMissing(Start) Then
        Min = 1
    Else
        Min = Start
    End If

    Do While Min <> Max - 1
        If MyXl.Worksheets(SheetNumber).Cells((Min + Max) / 2, 1) = "" Then
            Max = Int((Min + Max) / 2)
        Else
            Min = Int((Min + Max) / 2)
        End If
        Rem point = point + 1
    Loop
    NumMaxRows = Min

    MyXl.ActiveWorkbook.Saved = True
    MyXl.Quit
    Set MyXl = Nothing
End Function

Function fOSUserName() As String
' Returns the network login name
Dim lngLen As Long, lngX As Long
Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If lngX <> 0 Then
        fOSUserName = Left$(strUserName, lngLen - 1)
    Else
        fOSUserName = ""
    End If
End Function

Private Function fstrDField(mytext As String, delim As String, groupnum As Integer) As String

   ' this is a standard delimiter routine that every developer I know has.
   ' This routine has a million uses. This routine is great for splitting up
   ' data fields, or sending multiple parms to a openargs of a form
   '
   '  Parms are
   '        mytext   - a delimited string
   '        delim    - our delimiter (usually a , or / or a space)
   '        groupnum - which of the delimited values to return
   '

Dim startpos As Integer, endpos As Integer
Dim groupptr As Integer, chptr As Integer

chptr = 1
startpos = 0
 For groupptr = 1 To groupnum - 1
    chptr = InStr(chptr, mytext, delim)
    If chptr = 0 Then
       fstrDField = ""
       Exit Function
    Else
       chptr = chptr + 1
    End If
 Next groupptr
startpos = chptr
endpos = InStr(startpos + 1, mytext, delim)
If endpos = 0 Then
   endpos = Len(mytext) + 1
End If

fstrDField = Mid$(mytext, startpos, endpos - startpos)

End Function

Function GetDefaultPrinter() As String

   Dim strDefault    As String
   Dim lngbuf        As Long

   strDefault = String(255, Chr(0))
   lngbuf = GetProfileString("Windows", "Device", "", strDefault, Len(strDefault))
   If lngbuf > 0 Then
      GetDefaultPrinter = fstrDField(strDefault, ",", 1)
   Else
      GetDefaultPrinter = ""
   End If

End Function
Function CambiaStampantePredefinita(nomeStampante As String)
    Dim WshNet
    Set WshNet = CreateObject("Wscript.Network")
    WshNet.SetDefaultPrinter nomeStampante
    Set WshNet = Nothing
    CambiaStampantePredefinita = Null
End Function

Function Se_Aperto(ByVal NomeFile As String) As Boolean

    Dim Retval As Boolean
    Dim nfile As Integer

    On Local Error GoTo Se_ApertoErr

    nfile = FreeFile()
    Open NomeFile For Input Access Read Lock Read As nfile

Se_ApertoFine:
    Close nfile
    Se_Aperto = Retval
    Exit Function

Se_ApertoErr:
    If Err = 70 Then
       Retval = True
       MsgBox "Il file che si sta tentando di aprire è già in uso. Premere il tasto per riprovare", vbOK, "Attenzione"
    End If
Resume Se_ApertoFine

End Function


Sub CompattaBE()
   'Autore: Riccardo Pozzi e Federico Luciani

On Error GoTo gestErr
   Dim backend As String
   Dim strSQL As String
   Dim RS As DAO.Recordset
   Dim risp As Integer
   Dim copia As String
   Dim msg As String
   Dim nomebe As String
   Dim NumForms As Integer, I As Integer
   'Dim dbs As Database
   Dim tdf As TableDef

    NumForms = Forms.Count
'    Set dbs = CurrentDb()

    On Error GoTo ErrorHandler
    For I = NumForms - 1 To 0 Step -1
        'se si chiude una form durante il ciclo for, cambiano gli indici e possono
        'rimanere maschere aperte,
        'quindi bisogna cominciare dall'indice superiore fino a 0, per evitare
        'buchi nel ciclo

            DoCmd.Close acForm, Forms(I).Name, acSaveNo
    Next I

    strSQL = "SELECT Trim([Database]) AS DB " & _
            "FROM MSysObjects " & _
            "GROUP BY Trim([Database]), MSysObjects.Type " & _
            "HAVING (((MSysObjects.Type)=6));"
    Set RS = CurrentDb.OpenRecordset(strSQL)
    If RS.EOF And RS.BOF Then GoTo esci
    DoCmd.Hourglass True
    DoCmd.SetWarnings False
    RS.MoveFirst
    msg = ""
    backend = CurrentPath
    Kill (Mid(backend, 1, InStrRev(backend, "\")) & "a.new")
     DBEngine.CompactDatabase backend, Mid(backend, 1, InStrRev(backend, "\")) & "a.new"
    FileCopy Mid(backend, 1, InStrRev(backend, "\")) & "a.new", backend
    MsgBox "Compact operation completed" & _
      vbCrLf & msg, vbExclamation
esci:
    RS.Close
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Exit Sub
gestErr:
    MsgBox Err.Number & " - " & Err.Description
    GoTo esci
ErrorHandler:
Select Case Err.Number
    Case 53
        Resume Next
    Case 2501
    If Forms(I).Name = "maschera1" Then
        Resume Next
    End If
End Select
End Sub
Function GetNameCreditController(loginname As String) As String
Dim provv As Recordset
    'Set provv = New ADODB.Recordset
  '  With provv
'        .ActiveConnection = CurrentProject.Connection
 '       .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
  '      .MoveFirst
   '     .Find ("UserName='" & loginname & "'")
    '    If Not .EOF Then
     '       GetNameCreditController = .Fields("Name")
      '  Else
       '     GetNameCreditController = ""
        'End If
'        .Close
 '   End With
End Function
Function GetNameCreditControllerFromID(ID As Integer) As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   .Find ("ID=" & ID)
      '  If Not .EOF Then
       '     GetNameCreditControllerFromID = .Fields("Name")
        'Else
'            GetNameCreditControllerFromID = ""
 '       End If
  '      .Close
   ' End With
End Function

Function GetNumCreditController(loginname As String) As Integer

GetNumCreditController = DLookup("ID", "Tbl_Users", "'" & loginname & "'")
'Dim provv As Recordset
 '   Set provv = New ADODB.Recordset
  '  With provv
   '     .ActiveConnection = CurrentProject.Connection
    '    .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
     '   .MoveFirst
      '  .Find ("UserName='" & loginname & "'")
       ' If Not .EOF Then
        '    GetNumCreditController = .Fields("ID")
'        Else
 '           GetNumCreditController = 0
  '      End If
   '     .Close
    'End With
End Function
Function GetPahReleases() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPahReleases = .Fields("PathReleases")
      '  .Close
'    End With
End Function
Function GetImportingProcess() As Boolean
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    GetImportingProcess = .Fields("ImportingProcess")
     '   .Close
'    End With
End Function
Sub SetImportingProcess(Switch As Boolean)
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .Fields("ImportingProcess") = Switch
     '   .Update
      '  .Close
'    End With
End Sub


Function GetPahStatements() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPahStatements = .Fields("PathStatemets")
      '  .Close
'    End With
End Function

Function GetPahQueryFile() As String
Dim provv As Recordset
 '   Set provv = New ADODB.Recordset
  '  With provv
   '     .ActiveConnection = CurrentProject.Connection
    '    .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
     '   .MoveFirst
      '  GetPahQueryFile = .Fields("PathQueryFile")
       ' .Close
'    End With
End Function


Function GetLastUpdate() As Date
Dim provv As Recordset
 '   Set provv = New ADODB.Recordset
  '  With provv
   '     .ActiveConnection = CurrentProject.Connection
    '    .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
     '   .MoveFirst
      '  GetLastUpdate = .Fields("Lastupdate")
       ' .Close
'    End With
End Function

Function GetPathlogo() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPathlogo = .Fields("pathlogo")
      '  .Close
'    End With
End Function
Function GetPathImages() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPathImages = .Fields("PathImages")
      '  .Close
'    End With
End Function

Function GetPathExcelDirectory() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPathExcelDirectory = .Fields("PathExcelDirectory")
      '  .Close
'    End With
End Function

Function GetPathWordDirectory() As String
Dim provv As Recordset
'    Set provv = New ADODB.Recordset
 '   With provv
  '      .ActiveConnection = CurrentProject.Connection
   '     .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
    '    .MoveFirst
     '   GetPathWordDirectory = .Fields("PathWordDirectory")
      '  .Close
'    End With
End Function

Function NormalizeFileName(FileName As String) As String
    FileName = Replace(FileName, ".", "")
    FileName = Replace(FileName, Chr(34), "")
    FileName = Replace(FileName, "/", "")
    FileName = Replace(FileName, "\", "")
    FileName = Replace(FileName, "[", "")
    FileName = Replace(FileName, "]", "")
    FileName = Replace(FileName, ":", "")
    FileName = Replace(FileName, ";", "")
    FileName = Replace(FileName, "=", "")
    FileName = Replace(FileName, ",", "")
    NormalizeFileName = FileName
End Function

Function ExcelStatement(Customer As Recordset, CurrencyTab As Variant, rstbanks As Recordset, Optional Monthend As Date, Optional CloseStatement As Boolean) As String
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim ExeWksht As Excel.Worksheet
Dim Typecurrency, FileName As String
Dim NextEmptyColumn, sheet, Currinv, heet, c, r, row, col, StartingDataRow, StartingDataRow2, StartingDataRow3, InvoiceLine, COLWIDTH As Integer
Dim ColCustInvn, PullTicketN, OriginalAmount As Integer
Dim MonthEndAmount, Current, O31Days, O3160Days, O61days As Currency
Dim QuerySpec As Recordset
Dim DirSave As String
Dim aa, bb, ExtraColumns As Integer
Dim recc As DAO.Recordset
Dim DocumentTypes As Variant
Dim DocumentsToBeErased As Variant
Dim AddColumnCustomsInvoicenum, AddPullTicketn, AddOriginalTransactionAmount As Boolean

Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
'Set ExcDoc = ExcApp.Workbooks.Add
Dim PathTemplate As String
PathTemplate = DLookup("PathTemplates", "tblGeneral")
PathTemplate = PathTemplate & "\"
Set ExcDoc = ExcApp.Workbooks.Open(PathTemplate & "SimpleStatement.xlsx")
ExcApp.Visible = True
InvoiceLine = 1
sheet = 1
aa = 0

'conta quanti saranno le valute nell'estratto conto alla fine(tabs)
If CurrencyTab.Controls.item("Pagina88").Visible Then
    aa = aa + 1
End If
If CurrencyTab.Controls.item("Pagina89").Visible Then
    aa = aa + 1
End If
If CurrencyTab.Controls.item("Pagina90").Visible Then
    aa = aa + 1
End If
If CurrencyTab.Controls.item("Pagina91").Visible Then
    aa = aa + 1
End If
If CurrencyTab.Controls.item("Pagina92").Visible Then
    aa = aa + 1
End If

Set QuerySpec = New Recordset
With QuerySpec
    .ActiveConnection = CurrentProject.Connection
    .Open "Tbl_queries", , adOpenKeyset, adLockOptimistic, adCmdTable
End With

With rstbanks
    .MoveFirst
    .Find ("Country='" & Customer.Fields("Country") & "'")
End With

For bb = 1 To aa 'ripete per ogni tab (valuta)
    Select Case bb 'mette in recc la porzione di e/c corrispondente ad alla valuta che si sta esaminando
        Case 1
            Set recc = CurrencyTab.Controls.item("Pagina88").Controls.item(0).Form.RecordsetClone
        Case 2
            Set recc = CurrencyTab.Controls.item("Pagina89").Controls.item(0).Form.RecordsetClone
        Case 3
            Set recc = CurrencyTab.Controls.item("Pagina90").Controls.item(0).Form.RecordsetClone
        Case 4
            Set recc = CurrencyTab.Controls.item("Pagina91").Controls.item(0).Form.RecordsetClone
        Case 5
            Set recc = CurrencyTab.Controls.item("Pagina92").Controls.item(0).Form.RecordsetClone
    End Select
    recc.MoveFirst
    Currinv = 0

    Set DocumentTypes = New Recordset
    With DocumentTypes
        Set DocumentTypes = CurrentDb.OpenRecordset("SELECT Tbl_Types.ID, Tbl_Types.Descripition FROM Tbl_Types;")
    End With

    MonthEndAmount = 0
While (Not recc.EOF)
    AddColumnCustomsInvoicenum = Customer("PullTicketNumberToBePrinted")
    If AddColumnCustomsInvoicenum Then
        ColCustInvn = 2
    End If


    AddPullTicketn = Customer("FacturaNumberToBePrinted")
    If AddPullTicketn Then
        If AddColumnCustomsInvoicenum Then
            PullTicketN = 3
        Else
            PullTicketN = 2
        End If
    End If

    AddOriginalTransactionAmount = Customer("OriginalInvoiceAmountToBePrinted")
    If AddOriginalTransactionAmount Then
        OriginalAmount = 9
        If AddPullTicketn = False Then
            OriginalAmount = OriginalAmount - 1
        End If
        If AddColumnCustomsInvoicenum = False Then
            OriginalAmount = OriginalAmount - 1
        End If
    End If
    If ExcDoc.Worksheets.Count < sheet Then
        ExcDoc.Worksheets.Add
    End If
    ExcApp.Worksheets(sheet).Select
    row = 9
    col = 1
    With ExcApp
        Rem ExcApp.ActiveSheet.Insert.Picture (GetPathlogo())
Rem        .ActiveSheet.Pictures.Insert (GetPathlogo())
'        .ActiveSheet.Rows("1:1").RowHeight = 140

 '       .Cells(3, 1) = "STATEMENT OF ACCOUNT"
  '      .Cells(3, 1).Select
   '     .Cells(3, 1).Font.Bold = True
    '    .Cells(3, 1).Font.Italic = True
     '   .Cells(3, 1).Font.Size = 16
'        ExcDoc.Worksheets(sheet).Range("A3:G3").HorizontalAlignment = xlCenterAcrossSelection
 '       .Cells(row, col) = Customer.Fields("OWN_company")
  '      .Cells(row, col).Font.Bold = True
        'row = row + 1
        'r = row
        'c = col
        Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details1")
        '.Cells(row, col).Font.Bold = True
'        row = row + 1
        Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details2")
 '       .Cells(row, col).Font.Bold = True
  '      row = row + 1
        Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details3")
   '     .Cells(row, col).Font.Bold = True
    '    row = row + 1
        Rem .Cells(Row, Col) = customer.Fields("OWN_bank_details4")
     '   .Cells(row, col).Font.Bold = True
      '  row = row + 2
       ' col = 1
        .Cells(row, col) = Customer.Fields("Name")
        .Cells(row, col).Font.Bold = True
        .Cells(row, col + 4) = "Customer ID: " & Customer.Fields("customer_code")
        .Cells(row, col + 4).Font.Bold = True

        row = row + 1
        .Cells(row, col) = Customer.Fields("Address")
'        .Cells(row, col).Font.Bold = True
        row = row + 1
        .Cells(row, col) = Customer.Fields("Address2")
 '       .Cells(row, col).Font.Bold = True
        If Customer.Fields("Address3") <> "" Then
            row = row + 1
            .Cells(row, col) = Customer.Fields("Address3")
  '          .Cells(row, col).Font.Bold = True
        End If
        If Customer.Fields("Address4") <> "" Then
            row = row + 1
            .Cells(row, col) = Customer.Fields("Address4")
   '         .Cells(row, col).Font.Bold = True
        End If
        row = row + 1

        .Cells(row, col) = Customer.Fields("Country")

        .Cells(row, col + 4) = "Date: " & Format(Date, "dd-mmm-yyyy")
        .Cells(row, col + 4).Font.Bold = True

        row = row + 4
        StartingDataRow = row
        StartingDataRow2 = row
        StartingDataRow3 = row
'        .Cells(row, col + 0) = "Invoice number"
 '       .Cells(row, col + 1) = "Date"
  '      .Cells(row, col + 1).HorizontalAlignment = xlCenter

   '     .Cells(row, col + 2) = "Type"
    '    .Cells(row, col + 3) = "Reference"
     '   .Cells(row, col + 4) = "SO number"
      '  .Cells(row, col + 5) = "Amount due"
'        .Cells(row, col + 6).HorizontalAlignment = xlRight
 '       .Cells(row, col + 6) = "Due date"
  '      .Cells(row, col + 6).HorizontalAlignment = xlRight
        If Customer.Fields("StatementForm") > 0 Then
            If Customer.Fields("StatementForm") = 1 Then
                .Cells(row, col + 7) = "Queries"
                .Cells(row, col + 7).HorizontalAlignment = xlLeft
            ElseIf Customer.Fields("StatementForm") = 3 Then
                .Cells(row, col + 7) = "Queries"
                .Cells(row, col + 7).HorizontalAlignment = xlLeft
                .Cells(row, col + 8) = "Notes"
                .Cells(row, col + 8).HorizontalAlignment = xlLeft

            ElseIf Customer.Fields("StatementForm") = 4 Then
                .Cells(row, col + 7) = "Notes"
                .Cells(row, col + 7).HorizontalAlignment = xlLeft
            End If
        End If

        NextEmptyColumn = 1
        While .Cells(row, NextEmptyColumn) <> ""
            NextEmptyColumn = NextEmptyColumn + 1
        Wend
        If Monthend <> "00:00:00" Then
            .Cells(row, NextEmptyColumn) = "Total amount to be paid by month's end (" & Format(Monthend, "dd mmmm  yyyy") & ")"
            .Cells(row, NextEmptyColumn).HorizontalAlignment = xlCenter
            .Cells(row, NextEmptyColumn).WrapText = True
        End If

        'row = row + 1
        Typecurrency = recc("Tbl_Invoices.Currency")
        If Not (rstbanks.EOF) Then
            If Typecurrency = "EUR" Then
'                .Cells(r, c) = rstbanks.Fields("EURLine1")
 '               .Cells(r + 1, c) = rstbanks.Fields("EURLine2")
  '              .Cells(r + 2, c) = rstbanks.Fields("EURLine3")
   '             .Cells(r + 3, c) = rstbanks.Fields("EURLine4")
            ElseIf Typecurrency = "USD" Then
    '            .Cells(r, c) = rstbanks.Fields("USDLine1")
     '           .Cells(r + 1, c) = rstbanks.Fields("USDLine2")
      '          .Cells(r + 2, c) = rstbanks.Fields("USDLine3")
       '         .Cells(r + 3, c) = rstbanks.Fields("USDLine4")
            ElseIf Typecurrency = "GBP" Then
        '        .Cells(r, c) = rstbanks.Fields("GBPLine1")
         '       .Cells(r + 1, c) = rstbanks.Fields("GBPLine2")
          '      .Cells(r + 2, c) = rstbanks.Fields("GBPLine3")
           '     .Cells(r + 3, c) = rstbanks.Fields("GBPLine4")
            Else
            '    .Cells(r, c) = ""
             '   .Cells(r + 1, c) = ""
              '  .Cells(r + 2, c) = ""
               ' .Cells(r + 3, c) = ""
            End If
        End If
        Current = 0
        O31Days = 0
        O3160Days = 0
        O61days = 0
        InvoiceLine = row

        Set DocumentsToBeErased = New Recordset
        With DocumentsToBeErased
            Set DocumentsToBeErased = CurrentDb.OpenRecordset("SELECT Tbl_DocumentsToBeErased.CustomerID, Tbl_DocumentsToBeErased.DocumentType FROM Tbl_DocumentsToBeErased WHERE (((Tbl_DocumentsToBeErased.CustomerID)='" & Customer.Fields("customer_code") & "'));")
        End With

        On Error GoTo avanti
    While (Not recc.EOF)
        DocumentsToBeErased.FindFirst "Tbl_DocumentsToBeErased.DocumentType=" & ((recc.Fields("Type")))
        If (DocumentsToBeErased.RecordCount = 0) Or (DocumentsToBeErased.RecordCount <> 0 And DocumentsToBeErased.NoMatch = True) Then
            .Cells(row, col + 0) = recc.Fields("document_number")
            .Cells(row, col + 1) = Format(recc.Fields("date"), "dd/mmm/yyyy")
            DocumentTypes.MoveFirst
'            DocumentTypes.Move ((recc.Fields("Type"))) - 1
            '.Cells(row, col + 2) = DocumentTypes.Fields("Descripition")
            .Cells(row, col + 2) = recc.Fields("Customer_reference")
            .Cells(row, col + 3) = recc.Fields("SONumber")
            .Cells(row, col + 4) = recc.Fields("Tbl_Invoices.Currency")
            .Cells(row, col + 5) = Format(recc.Fields("amount"), "##,##0.00")
            .Cells(row, col + 6) = Format(recc.Fields("Tbl_Invoices.Overdue_Date"), "dd/mmm/yyyy")
            If Not IsNull(recc.Fields("Query")) Then
                QuerySpec.MoveFirst
                QuerySpec.Find ("ID=" & recc.Fields("Query"))
            End If
            If Customer.Fields("StatementForm") > 0 Then
                QuerySpec.MoveFirst
                If Not IsNull(recc.Fields("Query")) Then
                    QuerySpec.Find ("ID=" & recc.Fields("Query"))
                    If Not QuerySpec.EOF Then
                        If Customer.Fields("StatementForm") = 1 Then
                            If recc.Fields("QueryToBePrinted") = True Then
                                .Cells(row, col + 7) = QuerySpec.Fields("Query")
                            End If
                        ElseIf Customer.Fields("StatementForm") = 3 Then
                            If recc.Fields("QueryToBePrinted") = True Then
                                .Cells(row, col + 7) = QuerySpec.Fields("Query")
                            End If
                            If Not IsNull(recc.Fields("Memo")) Then
                                If recc.Fields("QueryToBePrinted") = True Then
                                    .Cells(row, col + 8) = recc.Fields("Memo")
                                End If
                            End If
                        ElseIf Customer.Fields("StatementForm") = 4 Then
                            If Not IsNull(recc.Fields("Memo")) Then
                                If recc.Fields("QueryToBePrinted") = True Then
                                    .Cells(row, col + 7) = recc.Fields("Memo")
                                End If
                            End If
                        End If
                        If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And _
                            ((QuerySpec.Fields("InvoiceToBePaid"))) Then
                                .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                                .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                        End If
                    Else
                        If recc.Fields("Tbl_Invoices.overdue_date") <= Monthend Then
                            .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                            .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                        End If
                    End If
                Else
                    If (Not (QuerySpec.EOF)) Then
                        'If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And (QuerySpec.Fields("InvoiceToBePaid")) Then
'                        If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) Then
 '                           .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
  '                          .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
   '                     End If
    '                Else
                        If (recc.Fields("Tbl_Invoices.Overdue_Date") <= Monthend) Then
                            .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                            .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                        End If
                    End If

                End If
                Else
                    If (Not (QuerySpec.EOF)) Then
                        'If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) And (QuerySpec.Fields("InvoiceToBePaid")) Then
'                        If (recc.Fields("Tbl_Invoices.overdue_date") <= Monthend) Then
 '                           .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
  '                          .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
   '                     End If
    '                Else
                        If (recc.Fields("Tbl_Invoices.Overdue_Date") <= Monthend) Then
                            .Cells(row, NextEmptyColumn) = Format(recc.Fields("amount"), "##,##0.00")
                            .Range("A" & Trim(Str(row)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row))).Interior.color = vbRed
                        End If
                    End If
                End If

                If recc.Fields("Tbl_Invoices.Overdue_Date") > Date Then
                    Current = Current + recc.Fields("amount")
                ElseIf DateAdd("d", 61, recc.Fields("Tbl_Invoices.Overdue_Date")) <= Date Then
                    O61days = O61days + recc.Fields("amount")
                ElseIf DateAdd("d", 31, recc.Fields("Tbl_Invoices.overdue_date")) <= Date Then
                    O3160Days = O3160Days + recc.Fields("amount")
                Else
                    O31Days = O31Days + recc.Fields("amount")

                End If
                recc.MoveNext
                row = row + 1
        Else
            recc.MoveNext
        End If
    Wend
avanti:
            row = row + 2
'            .Cells(row, 1) = "Balance due: " & Typecurrency
'            .Cells(row, 6).NumberFormat = "##,##0.00"
            ExcDoc.Worksheets(sheet).Cells(row, 4) = "Total Outstanding"
            ExcDoc.Worksheets(sheet).Cells(row, 6) = "=Sum(F1:F" & row - 1 & ")"
            'row = row + 2
            'ExcDoc.Worksheets(sheet).Cells(row, 4) = "Total overdue "


            row = row + 3
'            .Cells(row, 3) = "1-30 Days"
 '           .Cells(row, 3).Font.Bold = True
  '          .Cells(row, 3).HorizontalAlignment = xlCenter
   '         .Cells(row, 4) = "31-60 Days"
    '        .Cells(row, 4).Font.Bold = True
     '       .Cells(row, 4).HorizontalAlignment = xlCenter
      '      .Cells(row, 5) = "61+ Days"
       '     .Cells(row, 5).Font.Bold = True
        '    .Cells(row, 5).HorizontalAlignment = xlCenter
            row = row + 1
'            .Cells(row, 2) = "Current"
 '           .Cells(row, 2).Font.Bold = True
  '          .Cells(row, 2).HorizontalAlignment = xlCenter
   '         .Cells(row, 3) = "Past due"
    '        .Cells(row, 3).Font.Bold = True
     '       .Cells(row, 3).HorizontalAlignment = xlCenter
      '      .Cells(row, 4) = "Past due"
       '     .Cells(row, 4).Font.Bold = True
'            .Cells(row, 4).HorizontalAlignment = xlCenter
 '           .Cells(row, 5) = "Past due"
  '          .Cells(row, 5).Font.Bold = True
   '         .Cells(row, 5).HorizontalAlignment = xlCenter
            row = row + 2


'            .Cells(row, 1) = "Customer Bill-To Total:"
 '           .Cells(row, 1).Font.Bold = True
  '          .Cells(row, 2) = Format(Current, "##,##0.00")
   '         .Cells(row, 2).Font.Bold = True
    '        .Cells(row, 3) = Format(O31Days, "##,##0.00")
     '       .Cells(row, 3).Font.Bold = True
      '      .Cells(row, 4) = Format(O3160Days, "##,##0.00")
'            .Cells(row, 4).Font.Bold = True
 '           .Cells(row, 5) = Format(O61days, "##,##0.00")
  '          .Cells(row, 5).Font.Bold = True
'            ExcDoc.Worksheets(sheet).Range("A" & row - 6 & ":E" & row).Columns.AutoFit
 '           ExcDoc.Worksheets(sheet).Columns("b:c").AutoFit
  '          ExcDoc.Worksheets(sheet).Columns("e:f").AutoFit
'            If (Customer.Fields("StatementForm") = 3) Or (Customer.Fields("StatementForm") = 4) Then
 '               .Range("I:I").ColumnWidth = 50
  '              .Range("I:I").WrapText = True
   '         End If

'            If Monthend <> "00:00:00" Then
 '               .Cells(row + 2, 1) = "Total amount to be paid by month's end (" & Format(Monthend, "dd mmmm  yyyy") & ")"
  '              .Cells(row + 2, 1).Font.Bold = True
   '             .Cells(row + 2, NextEmptyColumn) = "=Sum(" & Chr(65 + NextEmptyColumn - 1) & "1:" & Chr(65 + NextEmptyColumn - 1) & row - 1 & ")"
    '            .Cells(row + 2, NextEmptyColumn).NumberFormat = "##,##0.00"
     '           .Range("A" & Trim(Str(row + 2)) & ":" & Chr(65 + NextEmptyColumn - 1) & Trim(Str(row + 2))).Interior.color = vbRed
      '          .Cells(row + 2, NextEmptyColumn).Font.Bold = True
       '         .Cells(row + 2, NextEmptyColumn).Columns.AutoFit
        '    End If

'            ExcDoc.Worksheets(sheet).Columns("g:h").AutoFit
 '           If ExcDoc.Worksheets(sheet).Columns("g").ColumnWidth < 16 Then
  '              ExcDoc.Worksheets(sheet).Columns("g").ColumnWidth = 16
   '         End If

Rem            .ActiveSheet.Shapes("Picture 1").Select

 '           With ExcDoc.Worksheets(sheet).PageSetup
'                .Zoom = False
 '               .FitToPagesWide = 1
  '              .FitToPagesTall = 1000
   '         End With

Rem            ExcApp.Selection.ShapeRange.IncrementLeft (-115)

'            If Customer.Fields("StatementForm") = 1 Then
 '               ExtraColumns = 1
  '          ElseIf Customer.Fields("StatementForm") = 3 Then
   '             ExtraColumns = 2
    '        End If
Rem            If .ActiveSheet.Shapes("Picture 1").Width < ExcDoc.Worksheets(sheet).Columns("A:" & Chr(Asc("G") + ExtraColumns)).Width Then
   Rem             ExcApp.Selection.ShapeRange.IncrementLeft (ExcDoc.Worksheets(sheet).Columns("A:" & Chr(Asc("G") + ExtraColumns)).Width - .ActiveSheet.Shapes("Picture 1").Width) / 2
      Rem      End If
Rem            ExcApp.Selection.ShapeRange.IncrementTop -5000
   Rem         ExcApp.Selection.ShapeRange.IncrementTop 5
            row = row
        row = row
    End With

'    ExcDoc.Worksheets(sheet).Name = Typecurrency

'    If AddColumnCustomsInvoicenum Then
 '       ExcDoc.Worksheets(sheet).Columns("B").Insert Shift:=xlToRight
 '       recc.MoveFirst
  '      ExcDoc.Worksheets(sheet).Cells(StartingDataRow, 2) = "Factura n#"
   '     StartingDataRow = StartingDataRow + 1
'        While Not recc.EOF
 '           ExcDoc.Worksheets(Sheet).Cells(StartingDataRow, 2) = recc.Fields("CustomsInvoiceNumber")
  '          ExcDoc.Worksheets(Sheet).Cells(StartingDataRow, 2).HorizontalAlignment = xlCenter
   '         StartingDataRow = StartingDataRow + 1
    '        recc.MoveNext
     '   Wend
'        ExcDoc.Worksheets(Sheet).Columns("B:B").AutoFit
'    End If

 '   If AddPullTicketn Then
  '      ExcDoc.Worksheets(sheet).Columns("B").Insert Shift:=xlToRight
'        recc.MoveFirst
   '     ExcDoc.Worksheets(sheet).Cells(StartingDataRow2, 2) = "Pull Ticket n#"
    'End If

'    If AddOriginalTransactionAmount Or AddPullTicketn Or AddColumnCustomsInvoicenum Then
 '       ExcDoc.Worksheets(sheet).Columns("" & Chr(64 + OriginalAmount) & "").Insert Shift:=xlToRight
  '      recc.MoveFirst
   '     ExcDoc.Worksheets(sheet).Cells(StartingDataRow3, OriginalAmount) = "Original Transaction Amount"
 '       StartingDataRow3 = StartingDataRow3 + 1
'        StartingDataRow = 18


'        While Not recc.EOF
 '           If AddColumnCustomsInvoicenum Then
  '              ExcDoc.Worksheets(sheet).Cells(StartingDataRow, ColCustInvn) = recc.Fields("CustomsInvoiceNumber")
   '             ExcDoc.Worksheets(sheet).Cells(StartingDataRow, ColCustInvn).HorizontalAlignment = xlCenter
    '        End If
     '
      '      If AddPullTicketn Then
       '         ExcDoc.Worksheets(sheet).Cells(StartingDataRow, PullTicketN) = recc.Fields("[PullTicketN#]")
        '        ExcDoc.Worksheets(sheet).Cells(StartingDataRow, PullTicketN).HorizontalAlignment = xlCenter
         '   End If

'            If AddOriginalTransactionAmount Then
 '               ExcDoc.Worksheets(sheet).Cells(StartingDataRow, OriginalAmount) = recc.Fields("[OriginalAmount]")
  '              ExcDoc.Worksheets(sheet).Cells(StartingDataRow, OriginalAmount).HorizontalAlignment = xlRight
   '         End If
    '        StartingDataRow = StartingDataRow + 1
     '       recc.MoveNext
      '  Wend
       ' ExcDoc.Worksheets(sheet).Columns("" & Chr(65 + col) & ":" & Chr(65 + col) & "").AutoFit
'    End If
 '   sheet = sheet + 1
Wend
Next bb
row = row

'Call AddPaymentsReceived(ExcApp, ExcDoc, Customer.Fields("Customer_code"))
'ExcApp.ActiveSheet.PageSetup.CenterFooter = "Pag. &P of &N"

DirSave = DLookup("PathStatemets", "TblGeneral")
DirSave = Replace(DirSave, "*username*", fOSUserName())
'If Dir(Dirsave, 16) = "" Then
 '   MkDir (Dirsave)
'End If

'Dirsave = "C:\Users\" & fOSUserName() & "\Statements\" & Customer.Fields("Country") & "\"
'If Dir(Dirsave, 16) = "" Then
 '   MkDir (Dirsave)
'End If

'Dirsave = "C:\Users\" & fOSUserName() & "\Statements\" & Customer.Fields("Country") & "\" & Format((Date), "dd-mm-yyyy") & "\"
'If Dir(Dirsave, 16) = "" Then
 '   MkDir (Dirsave)
'End If

FileName = Customer.Fields("Name")
FileName = NormalizeFileName(FileName)

ExcelStatement = DirSave & FileName & " - " & Format((Now), "dd mmm yyyy - hh.mm.ss") & ".xlsx"
ExcApp.ActiveWorkbook.SaveAs FileName:=ExcelStatement, FileFormat:=xlWorkbookDefault
If CloseStatement = True Then
    ExcApp.Quit
    Set ExcApp = Nothing
    Set ExcDoc = Nothing
    DocumentTypes.Close
    DocumentTypes = Null
    DocumentsToBeErased.Close
End If
QuerySpec.Close
End Function

Sub UpdateQueryFile(CustomerRst As Variant)
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim Riga, Column, I As Integer
Dim Found As Boolean
Dim rst, InvoicesRst, a As Recordset
Dim RstAdditionalQueryData As Variant
Dim rngXL As Excel.Range

Dim S As String
GoTo fine
    Set ExcApp = CreateObject("Excel.Application")
    Rem Set ExcDoc = ExcApp.Workbooks.Open(GetPahQueryFile & GetNameCreditController(fOSUserName) & ".xls")
Rem    Set ExcDoc = ExcApp.Workbooks.Open(GetPahQueryFile & GetNameCreditControllerFromID(CustomerRst.Credit_controller) & ".xls")

    ExcApp.Visible = True
    Found = False
    Rem questo ciclo va ripetuto per ciascuna riga con la causale query <> nil
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Queries", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
    End With

    Set InvoicesRst = New ADODB.Recordset
    With InvoicesRst
        Set InvoicesRst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.*, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date, Tbl_Invoices.Query FROM Tbl_Invoices WHERE (((Tbl_Invoices.Customer_ID)=" & CustomerRst.Fields("Customer_code") & ") AND ((Tbl_Invoices.Update_date)=#" & Format(Now(), "mm/dd/yyyy") & "#) AND ((Tbl_Invoices.Query)>0));")
    End With

    Set RstAdditionalQueryData = CurrentDb.OpenRecordset("SELECT Tbl_AdditionalQueryData.* FROM Tbl_AdditionalQueryData;")

    While Not InvoicesRst.EOF
        If InvoicesRst.Fields("Tbl_Invoices.Query") > 0 Then
            Riga = 2
            Found = False
            With ExcDoc
                While (.Worksheets(1).Cells(Riga, 1) <> "") And (Found = False)
                    'aggiorana queryfile.xls
                    If .Worksheets(1).Cells(Riga, 4) = CustomerRst.Fields("Customer_code") Then
                        If .Worksheets(1).Cells(Riga, 5) = InvoicesRst.Fields("Document_Number") Then
                            Found = True
                            Riga = Riga - 1
                        ElseIf .Worksheets(1).Cells(Riga, 4) <> .Worksheets(1).Cells(Riga + 1, 4) Then
                            Found = True
                            Set rngXL = ExcDoc.Sheets(1).Rows(Riga)
                            rngXL.Insert xlShiftDown
                            Rem rngXL = Nothing
                            Rem .Worksheets(1).Range(Cells(Riga, 1), Cells(Riga, 1)).EntireRow.Insert
                            .Worksheets(1).Cells(Riga, 1) = Date
                            Riga = Riga - 1

                            RstAdditionalQueryData.MoveFirst
                            S = "Customer_Code=" & InvoicesRst.Fields("Tbl_Invoices.Customer_ID") & " AND Document_Number= '" & InvoicesRst.Fields("document_number") & "'" & " AND Query_Date=#" & Format(InvoicesRst.Fields("Date"), "mm/dd/yy") & "#"
                            RstAdditionalQueryData.FindFirst S
                            If RstAdditionalQueryData.NoMatch = True Then
                                RstAdditionalQueryData.AddNew
                            Else
                                RstAdditionalQueryData.Edit
                            End If
                            RstAdditionalQueryData.Fields("Customer_Code") = InvoicesRst.Fields("Tbl_Invoices.Customer_ID")
                            RstAdditionalQueryData.Fields("Document_Number") = InvoicesRst.Fields("document_number")
                            RstAdditionalQueryData.Fields("Document_date") = InvoicesRst.Fields("Date")
                            RstAdditionalQueryData.Fields("Query_date") = Date
                        Rem    RstAdditionalQueryData.Update

                        End If
                    End If
                    Riga = Riga + 1
                Wend
                If .Worksheets(1).Cells(Riga, 1) = "" Then
                    .Worksheets(1).Cells(Riga, 1) = Date
                End If
                .Worksheets(1).Cells(Riga, 2) = CustomerRst.Fields("country")
                .Worksheets(1).Cells(Riga, 3) = CustomerRst.Fields("Tbl_Customers.Name")
                .Worksheets(1).Cells(Riga, 4) = InvoicesRst.Fields("Tbl_Invoices.Customer_ID")
                .Worksheets(1).Cells(Riga, 6) = InvoicesRst.Fields("Date")
                .Worksheets(1).Cells(Riga, 5) = InvoicesRst.Fields("document_number")
                .Worksheets(1).Cells(Riga, 7) = InvoicesRst.Fields("Overdue_Date")
                .Worksheets(1).Cells(Riga, 8) = InvoicesRst.Fields("amount")
                .Worksheets(1).Cells(Riga, 9) = InvoicesRst.Fields("Currency")
                .Worksheets(1).Cells(Riga, 12) = InvoicesRst.Fields("Memo")
                rst.MoveFirst
                rst.Find ("ID='" & InvoicesRst.Fields("Tbl_Invoices.Query")) & "'"
                If Not (rst.EOF) Then
                    .Worksheets(1).Cells(Riga, 11) = rst.Fields("query")
                    .Worksheets(1).Cells(Riga, 10) = rst.Fields("Resolution_owner")
                End If

Rem
            End With
        End If
        InvoicesRst.MoveNext
    Wend
    ExcDoc.Save
    ExcDoc.Close
    Set ExcDoc = Nothing
    Set ExcApp = Nothing

    Rem (wdDoNotSaveChanges)
    rst.Close
    InvoicesRst.Close
    RstAdditionalQueryData.Close
    QueryFileTochange = False
    ChargebackFileTochange = False

    Set ExcDoc = Nothing
    Set ExcApp = Nothing
    Set rst = Nothing
    Set RstAdditionalQueryData = Nothing
fine:
End Sub
Function FindLastDate() As Date

Dim a, rst, b As Recordset
FindLastDate = DMax("Update_date", "Tbl_Invoices")
'    With rst
 '       Set rst = New ADODB.Recordset
  '      Set rst = CurrentDb.OpenRecordset("SELECT Max(Tbl_Invoices.Update_date) AS MaxOfUpdate_date FROM Tbl_Invoices;")
   '     FindLastDate = .Fields("MaxOfUpdate_date")

Rem        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date FROM Tbl_Invoices GROUP BY Tbl_Invoices.Update_date;")
   Rem     .MoveLast
      Rem  FindLastDate = .Fields("Update_date")
    '    .Close
    'End With
End Function

Function FindPreviousDate() As Date
Dim a, rst, b As Recordset
    With rst
        Set rst = New ADODB.Recordset
        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.Update_date FROM Tbl_Invoices GROUP BY Tbl_Invoices.Update_date;")
        .MoveLast
        .MovePrevious
        If rst.RecordCount > 1 Then
            FindPreviousDate = .Fields("Update_date")
        Else
            FindPreviousDate = DateAdd("d", -1, Now())
        End If
        .Close
    End With
End Function
Sub QueryClosed(RstCustomer, RstInvoices As Variant)
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim row As Integer
Dim ExitLoop As Boolean

    Set ExcApp = CreateObject("Excel.Application")
    On Error GoTo CloseFiles
    Set ExcDoc = ExcApp.Workbooks.Open(GetPahQueryFile & GetNameCreditControllerFromID(RstCustomer.Fields("Credit_controller")) & ".xls")
    ExcApp.Visible = False
    row = 2
    With ExcDoc.Sheets(1)
        ExitLoop = False
        While ExitLoop = False
            If (.Cells(row, 4) = RstCustomer.Fields("Customer_code")) And (.Cells(row, 5) = RstInvoices.Fields("Document_Number")) Then
                .Cells(row, 14) = Date
                .Cells(row, 15) = "Invoice closed"
                ExitLoop = True
            ElseIf .Cells(row, 4) = "" Then
                ExitLoop = True
            Else
                row = row + 1
            End If
        Wend
    End With
    ExcDoc.Save
    ExcDoc.Close
    GoTo CloseFiles

CloseFiles:
    Set ExcDoc = Nothing
    Set ExcApp = Nothing
End Sub

Function FirstDayMonth() As Date
    FirstDayMonth = DateAdd("d", -Day(DateAdd("m", 1, Date)), (DateAdd("m", 1, Date)))
End Function

Public Sub Totals(CurrencyName As String, CustomerID As Integer, Maschera1 As Variant)

End Sub
Public Sub FindCustomerLastDate()
    CustomerLastDate = FindLastDate
End Sub

Public Function CurrentPath() As String
    'Restituisce: il percorso completo del database corrente (BE-FE)
    Dim tab_az As String
    tab_az = CurrentDb.TableDefs("Tbl_Customers").Connect
    CurrentPath = Mid(tab_az, 11, Len(tab_az))
End Function

Sub ShowUserRosterMultipleUsers()
    Dim cn As New ADODB.Connection
    Dim cn2 As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim I, j As Long
    Dim CompactDatabase As Boolean

    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open "Data Source=" & CurrentPath

    Rem cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
    & "Data Source=Q:\Credit control\Access\db1.mdb"

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set RS = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.

    Rem Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, _
    "", rs.Fields(2).Name, rs.Fields(3).Name

    CompactDatabase = True
    While Not RS.EOF
        If Interaction.Environ("Computername") <> Trim(RS.Fields(0)) Then
            CompactDatabase = False
        End If
        Rem Debug.Print rs.Fields(0), rs.Fields(1), _
        rem rs.Fields(2), rs.Fields(3)
        RS.MoveNext
    Wend
    cn.Close
    If CompactDatabase = True Then
        Call CompattaBE
    End If


End Sub
Function PrintQueryWithoutCreditController() As Boolean
Dim provv As Recordset
    Set provv = New ADODB.Recordset
    With provv
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
        .Find ("Username='" & fOSUserName() & "'")
        PrintQueryWithoutCreditController = False
        If (.Fields("Querywithoutcreditcontroller") = True) And Not (IsNull(.Fields("QuerywithoutcreditcontrollerEvery"))) Then
            PrintQueryWithoutCreditController = (.Fields("Querywithoutcreditcontroller") = True) And (CDate(CLng(Right(.Fields("QuerywithoutcreditcontrollerEvery"), 5))) <= Now())
        End If
        .Close
    End With
End Function

Function PrintQueryOnAccounts() As Boolean
Dim provv As Recordset
    Set provv = New ADODB.Recordset
    With provv
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        PrintQueryOnAccounts = False
        Rem.Find ("Customer_code=" & Testo10.Value)
        If (.Fields("Onaccountsstillopen") = True) And Not (IsNull(.Fields("OnaccountsstillopenEvery"))) Then
            PrintQueryOnAccounts = (.Fields("Onaccountsstillopen") = True) And (CDate(CLng(Right(.Fields("OnaccountsstillopenEvery"), 5))) <= Now())
        End If
        .Close
    End With
End Function

Function TotalInvoicesSelected(rst As Variant, SelTop, SelHeight As Long) As Currency
    Dim I As Long
    If rst.RecordCount > 0 Then
        TotalInvoicesSelected = 0
        rst.MoveFirst
        rst.Move (SelTop - 1)
        For I = 1 To SelHeight
            TotalInvoicesSelected = TotalInvoicesSelected + rst.Fields("amount")
            rst.MoveNext
        Next I
    End If
End Function

Function ExtractDateTimezone(S As String) As Date
    Dim WDay, NMonth, NSunday As Integer
    If Left(S, 1) = "D" Then
        Rem data precisa
        ExtractDateTimezone = DateSerial(Year(Now()), CInt(Mid(S, 4, 2)), CInt(Mid(S, 2, 2)))
    Else
        Rem numero domenica
        NSunday = CInt(Mid(S, 2, 2))
        NMonth = CInt(Mid(S, 4, 2))
        ExtractDateTimezone = DateSerial(Year(Now()), NMonth, 1)
        If NSunday < 99 Then
            WDay = Weekday(ExtractDateTimezone)
            If WDay <> 1 Then
                ExtractDateTimezone = DateAdd("d", 8 - WDay, ExtractDateTimezone)
            End If
        Else
            ExtractDateTimezone = DateAdd("m", 1, ExtractDateTimezone)
            WDay = Weekday(ExtractDateTimezone)
            If WDay <> 1 Then
                ExtractDateTimezone = DateAdd("d", 8 - WDay, ExtractDateTimezone)
            End If
            ExtractDateTimezone = DateAdd("d", -7, ExtractDateTimezone)
        End If
    End If
End Function

Function GetNextMonthEnd() As Variant
GetNextMonthEnd = DMin("MonthEnd", "Tbl_MonthEnd", "[MonthEnd]>=#" & Date & "#")
If IsNull(GetNextMonthEnd) Then GetNextMonthEnd = DMax("MonthEnd", "Tbl_MonthEnd")
'Dim provv As Recordset
'Dim a As Integer
 '   Set provv = New ADODB.Recordset
  '  With provv
   '     .ActiveConnection = CurrentProject.Connection
    '    .Open "Tbl_MonthEnd", , adOpenKeyset, adLockOptimistic, adCmdTable
     '   .MoveFirst
      '  While Not .EOF
       '     If .Fields("MonthEnd") >= Date Then
        '        If (GetNextMonthEnd >= .Fields("MonthEnd")) Or (GetNextMonthEnd = "") Then
         '           GetNextMonthEnd = CDate(.Fields("MonthEnd"))
          '      End If
           ' End If
            '.MoveNext
'        Wend
 '       .Close
  '  End With
End Function
Function GetLatestPaymentDate() As Variant
Dim provv As Recordset
Dim a As Integer
    Set provv = New ADODB.Recordset
    With provv
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_MonthEnd", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        While Not .EOF
            If .Fields("LatestPaymentDate") >= Date Then
                If (GetLatestPaymentDate >= .Fields("LatestPaymentDate")) Or (GetLatestPaymentDate = "") Then
                    GetLatestPaymentDate = CDate(.Fields("LatestPaymentDate"))
                End If
            End If
            .MoveNext
        Wend
        .Close
    End With
End Function


Function TemplateBitsCount(S As String) As Integer
Dim I As Integer
    If Len(S) > 0 Then
        TemplateBitsCount = 1
        For I = 1 To Len(S)
            If Mid$(S, I, 1) = "}" Then
                TemplateBitsCount = TemplateBitsCount + 1
                If Mid$(S, I + 1, 1) <> "{" And Len(S) > I Then
                    TemplateBitsCount = TemplateBitsCount + 1
                End If
            End If
        Next I
    Else
        TemplateBitsCount = 0
    End If
End Function

Function DivideTemplateInBits(S As String, NBits As Integer, Customer As Variant, Optional RecipientsName As String) As String
Dim Sentences()
Dim POS, I, EndSentence As Integer
Dim RS As DAO.Recordset
Dim Currenciess As String
Dim tot As Currency
    ReDim Sentences(NBits)
    POS = 1
    For I = 1 To Len(S)
        If Mid$(S, I, 1) = "{" Then
            EndSentence = InStr(I, S, "}")
            Sentences(POS) = Mid$(S, I, EndSentence - I + 1)
            I = EndSentence + 1
        Else
            EndSentence = InStr(I, S, "{") - 1
            If EndSentence < 1 Then
                EndSentence = Len(S)
            End If
            Sentences(POS) = Mid$(S, I, EndSentence)
            I = InStr(I, S, "{") - 1
            If POS = NBits Then
                I = Len(S)
            End If
        End If
        If Left(Sentences(POS), 1) = "{" Then
            Sentences(POS) = Mid(Sentences(POS), 2, Len(Sentences(POS)) - 1)
        End If
        If Right(Sentences(POS), 1) = "}" Then
            Sentences(POS) = Mid(Sentences(POS), 1, Len(Sentences(POS)) - 1)
        End If
        POS = POS + 1
    Next I
    Rem sostituisce delimitatori a dati
    For I = 1 To UBound(Sentences)
        While InStr(1, Sentences(I), "«") <> 0
            S = Mid$(Sentences(I), InStr(1, Sentences(I), "«") + 1, (InStr(1, Sentences(I), "»") - 1) - InStr(1, Sentences(I), "«"))
            Select Case S
                Case "1": Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Customer.Fields("ContactNames").value & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "2": Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Customer.Fields("Tbl_Customers.Name").value & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "3": Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Format(GetNextMonthEnd, "dd-mm-yyyy") & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "4":
                    Set RS = CurrentDb.OpenRecordset("SELECT Sum(Tbl_Invoices.Amount) AS SommaDiAmount, Tbl_Invoices.Customer_ID, Tbl_Invoices.Currency FROM Tbl_queries RIGHT JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query WHERE (((Tbl_Invoices.Update_date) = Date()) And ((Tbl_Invoices.Overdue_Date) <= #" & Format(GetNextMonthEnd, "mm/dd/yy") & "#) And ((Tbl_queries.InvoiceToBePaid) = Yes Or (Tbl_queries.InvoiceToBePaid) Is Null)) GROUP BY Tbl_Invoices.Customer_ID, Tbl_Invoices.Currency HAVING (((Tbl_Invoices.Customer_ID)=" & Customer.Fields("Customer_code").value & "));")
                    Currenciess = ""
                    If RS.RecordCount > 0 Then
                        RS.MoveFirst
                        tot = 0
                        While Not RS.EOF
                            Currenciess = Currenciess & RS.Fields("currency").value & " " & Format(RS.Fields("SommaDiAmount").value, "##,##0.00") & " + "
                            tot = tot + RS.Fields("SommaDiAmount").value * DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='" & RS.Fields("Currency").value & "'")
                            RS.MoveNext
                        Wend
                    End If
                    If Len(Currenciess) > 0 Then
                        Currenciess = Left(Currenciess, Len(Currenciess) - 2)
                    Else
                        Currenciess = "0"
                    End If
                    If tot > 0 Then
                        Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Currenciess & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                    Else
                        Sentences(I) = ""
                    End If
                Case "5":
                    Set RS = CurrentDb.OpenRecordset("SELECT Sum(Tbl_Invoices.Amount) AS SommaDiAmount, Tbl_Invoices.Customer_ID, Tbl_Invoices.Currency FROM Tbl_queries RIGHT JOIN Tbl_Invoices ON Tbl_queries.ID = Tbl_Invoices.Query WHERE (((Tbl_Invoices.Update_date) = Date()) And ((Tbl_Invoices.Overdue_Date) <= #" & Format(Date, "mm/dd/yy") & "#) And ((Tbl_queries.InvoiceToBePaid) = Yes Or (Tbl_queries.InvoiceToBePaid) Is Null)) GROUP BY Tbl_Invoices.Customer_ID, Tbl_Invoices.Currency HAVING (((Tbl_Invoices.Customer_ID)=" & Customer.Fields("Customer_code").value & "));")
                    Currenciess = ""
                    If RS.RecordCount > 0 Then
                        RS.MoveFirst
                        tot = 0
                        While Not RS.EOF
                            Currenciess = Currenciess & RS.Fields("currency").value & " " & Format(RS.Fields("SommaDiAmount").value, "##,##0.00") & " + "
                            Rem tot = tot + rs.Fields("SommaDiAmount").Value
                            tot = tot + RS.Fields("SommaDiAmount").value * Nz(RS.Fields("SommaDiAmount").value * DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='" & RS.Fields("Currency").value & "'"), 0)
                            RS.MoveNext
                        Wend
                    End If
                    If Len(Currenciess) > 0 Then
                        Currenciess = Left(Currenciess, Len(Currenciess) - 2)
                    Else
                        Currenciess = "0"
                    End If

                    If tot > 0 Then
                        Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Currenciess & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                    Else
                        Currenciess = ""
                        Rem Sentences(i) = Left$(Sentences(i), InStr(1, Sentences(i), "«") - 1) & Currenciess & Right$(Sentences(i), Len(Sentences(i)) - InStr(1, Sentences(i), "»"))
                        Sentences(I) = ""
                    End If
                Case "6": Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Customer.Fields("TotalInsurance").value & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "7": Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Customer.Fields("Description").value & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "10":
                    Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & GetCreditControllerSignature(Customer.Fields("Credit_controller")) & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "11":
                    Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Customer.Fields("Customer_code") & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "12":
                    Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Format(Now(), "dd-mm-yy") & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "13":
                    Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & GetNameCreditController(fOSUserName) & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                Case "14":
                    If IsNull(RecipientsName) Then
                        Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & "" & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                    Else
                        Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & RecipientsName & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
                    End If
                Case "15":
                    Sentences(I) = Left$(Sentences(I), InStr(1, Sentences(I), "«") - 1) & Format(GetLatestPaymentDate, "dd-mm-yyyy") & Right$(Sentences(I), Len(Sentences(I)) - InStr(1, Sentences(I), "»"))
            End Select

        Wend
    Next I
    For I = 1 To UBound(Sentences)
        If Sentences(I) <> "" Then
            DivideTemplateInBits = DivideTemplateInBits & Sentences(I)
        End If
    Next I
    Rem Set rs = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.*, Tbl_Types.Descripition, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date FROM Tbl_Types INNER JOIN Tbl_Invoices ON Tbl_Types.ID = Tbl_Invoices.Type WHERE (((Tbl_Invoices.Customer_ID)=" & Testo10.Value & ") AND ((Tbl_Invoices.Update_date)=#" & Format(Now(), "mm/dd/yyyy") & "#)) ORDER BY Tbl_Invoices.Currency,Tbl_Invoices.Overdue_Date, Tbl_Invoices.Document_Number;")
End Function

Function GetCreditControllerSignature(ID As Integer) As String
Dim provv As Recordset
    Set provv = New ADODB.Recordset
    With provv
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_Users", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        .Find ("ID=" & ID)
        If Not .EOF Then
            GetCreditControllerSignature = .Fields("Signature")
        Else
            GetCreditControllerSignature = ""
        End If
        .Close
    End With
End Function

Function GetPathChargbackFile() As String
Dim provv As Recordset
    Set provv = New ADODB.Recordset
    With provv
        .ActiveConnection = CurrentProject.Connection
        .Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        GetPathChargbackFile = .Fields("PathChargebackFile")
        .Close
    End With
End Function

Function FoundInExcelFile(XLSFile As Variant, ParamArray Options()) As Integer
Dim Line, I As Integer
Dim Found As Boolean
    FoundInExcelFile = 0
    Line = Options(0)
    Found = False
    While (XLSFile.Cells(Line, 1) <> "") And (Found = False)
        Found = True
        For I = 1 To UBound(Options) Step 2
            If Found = True Then
                Found = XLSFile.Cells(Line, Options(I)) = Options(I + 1)
            End If
        Next I
        If Found = True Then
            FoundInExcelFile = Line
        Else
            Line = Line + 1
        End If
    Wend
End Function

Sub UpdateChargebackFile(CustomerRst As Variant)
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim row, Riga, Column As Integer
Dim rst, S As Variant

    Set rst = New ADODB.Recordset
    With rst
        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Invoices.*, Tbl_Types.ToFillChargbackFile, Tbl_Invoices.Update_date, Tbl_Invoices.mEMO, Tbl_Invoices.Customer_ID FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Types.ToFillChargbackFile)=True) AND ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#) AND ((Tbl_Invoices.mEMO) Is Not Null) AND ((Tbl_Invoices.Customer_ID)=" & CustomerRst.Fields("Customer_code") & "));")
        If .RecordCount > 0 Then
            Set ExcApp = CreateObject("Excel.Application")
            Set ExcDoc = ExcApp.Workbooks.Open(GetPathChargbackFile & "\chargebacks.xls")
            ExcApp.Visible = True
            row = NumMaxRows(GetPathChargbackFile & "\chargebacks.xls", "", 7)
            .MoveLast
            .MoveFirst
            While Not rst.EOF
                Riga = 7
                With ExcDoc
                    While CDbl(.Worksheets(1).Cells(Riga, 1) < CDbl(rst.Fields("Document_Number")))
                        Riga = Riga + 1
                    Wend
                    If (CDbl(.Worksheets(1).Cells(Riga, 1)) = CDbl(rst.Fields("Document_number"))) Then
                        .Worksheets(1).Cells(Riga, 4) = rst.Fields("Tbl_Invoices.Memo")
                    End If
                End With
                rst.MoveNext
            Wend
            ExcDoc.Save
            ExcDoc.Close
            Set ExcDoc = Nothing
            Set ExcApp = Nothing
        End If
    End With
    rst.Close
    Set rst = Nothing
End Sub
Sub OLDUpdateChargebackFileAfterImport()
Dim rst As Variant
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim row, TestRow, I As Integer
Dim FoundChargeback As Boolean

    CurrentDb.QueryDefs("QueryChargebacksOpenNow").SQL = "SELECT Tbl_Invoices.*, Tbl_Types.ToFillChargbackFile, Tbl_Invoices.Update_date FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Types.ToFillChargbackFile)=True) AND ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#));"
    CurrentDb.QueryDefs("QueryChargebacksOpenPreviousDate").SQL = "SELECT Tbl_Invoices.*, Tbl_Types.ToFillChargbackFile, Tbl_Invoices.Update_date FROM Tbl_Invoices INNER JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_Types.ToFillChargbackFile)=True) AND ((Tbl_Invoices.Update_date)=#" & Format(FindPreviousDate, "mm/dd/yy") & "#));"

    Rem newchargebacks
    Set rst = CurrentDb.OpenRecordset("SELECT QueryChargebacksOpenNow.Customer_ID, Tbl_Customers.Name, QueryChargebacksOpenPreviousDate.Document_Number, QueryChargebacksOpenPreviousDate.*,Tbl_Customers.* FROM (QueryChargebacksOpenPreviousDate LEFT JOIN QueryChargebacksOpenNow ON (QueryChargebacksOpenPreviousDate.Customer_ID = QueryChargebacksOpenNow.Customer_ID) AND (QueryChargebacksOpenPreviousDate.Document_Number = QueryChargebacksOpenNow.Document_Number)) INNER JOIN Tbl_Customers ON QueryChargebacksOpenPreviousDate.Customer_ID = Tbl_Customers.Customer_code WHERE (((QueryChargebacksOpenNow.Customer_ID) Is not Null)) ORDER BY QueryChargebacksOpenPreviousDate.Document_Number;")

    Set ExcApp = CreateObject("Excel.Application")
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathChargbackFile & "\chargebacks.xls")
    ExcApp.Visible = False

    Rem If there is a new chargeback, Access adds a line to the file
    If rst.RecordCount > 0 Then
        row = NumMaxRows(GetPathChargbackFile & "\chargebacks.xls", "", 7) + 1
        TestRow = row
        rst.MoveFirst
        While Not rst.EOF
            With ExcDoc.Sheets(1)
                FoundChargeback = False
                For I = 7 To TestRow
                    If (.Cells(I, 3) = rst.Fields("QueryChargebacksOpenNow.Customer_ID")) And _
                    (.Cells(I, 7) = rst.Fields("Date")) And _
                    (.Cells(I, 1) = CCur(rst.Fields("QueryChargebacksOpenPreviousDate.Document_Number"))) Then
                        FoundChargeback = True
                        I = TestRow
                    End If
                Next I

                If FoundChargeback = False Then
                    .Cells(TestRow, 3) = rst.Fields("QueryChargebacksOpenNow.Customer_ID")
                    .Cells(TestRow, 7) = rst.Fields("Date")
                    .Cells(TestRow, 1) = CDbl(rst.Fields("QueryChargebacksOpenPreviousDate.Document_Number"))
                    .Cells(TestRow, 6) = rst.Fields("Amount")
                    .Cells(TestRow, 5) = rst.Fields("Currency")
                    .Cells(TestRow, 2) = rst.Fields("Tbl_Customers.Name")
                    .Cells(TestRow, 8) = GetNameCreditControllerFromID(rst.Fields("credit_controller"))
                    TestRow = TestRow + 1
                End If
            End With
            rst.MoveNext
        Wend
        ExcDoc.Save
    End If
    Set rst = Nothing

    Rem If a chargeback has been close Access puts the closure date
    Set rst = CurrentDb.OpenRecordset("SELECT QueryChargebacksOpenNow.Customer_ID, Tbl_Customers.Name, QueryChargebacksOpenPreviousDate.Document_Number, QueryChargebacksOpenPreviousDate.* FROM (QueryChargebacksOpenPreviousDate LEFT JOIN QueryChargebacksOpenNow ON (QueryChargebacksOpenPreviousDate.Customer_ID = QueryChargebacksOpenNow.Customer_ID) AND (QueryChargebacksOpenPreviousDate.Document_Number = QueryChargebacksOpenNow.Document_Number)) INNER JOIN Tbl_Customers ON QueryChargebacksOpenPreviousDate.Customer_ID = Tbl_Customers.Customer_code WHERE (((QueryChargebacksOpenNow.Customer_ID) Is Null)) ORDER BY QueryChargebacksOpenPreviousDate.Document_Number;")
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        While Not rst.EOF
            With ExcDoc.Sheets(1)
                row = FoundInExcelFile(ExcDoc.Sheets(1), 7, 1, rst.Fields("QueryChargebacksOpenPreviousDate.Document_Number"), 3, rst.Fields("QueryChargebacksOpenPreviousDate.Customer_ID"))
                Rem put closing date in Excel file
                If row <> 0 Then
                    .Cells(row, 9) = Format(Date, "dd-mmm-yy")
                End If
            End With
            rst.MoveNext
        Wend
        ExcDoc.Save
    End If
    ExcDoc.Close
    ExcApp.Close
    Set ExcApp = Nothing
    Set ExcDoc = Nothing
    Set rst = Nothing
    DoEvents
End Sub
Function GetApproverEmailAddress(CLExcess As Currency, StringTo, StringCC As String) As String
Dim rst As Recordset
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "TbReleasesApprovalMatrix", , adOpenKeyset, adLockOptimistic, adCmdTable
        .MoveFirst
        StringCC = ""
        While .Fields("ApprovalLimit") < CLExcess
            StringCC = StringCC & .Fields("EmailAddress") & ","
            .MoveNext
        Wend
        StringTo = .Fields("EmailAddress")
        GetApproverEmailAddress = .Fields("Name")
        .Close
    End With
End Function
Sub CopyStatement()
    Dim rst, rst2 As Variant
    Dim NextFiscalMonthEnd, PrevMonthEnd As Date
    Dim I As Integer

    If DCount("Update_date", "Tbl_Historical_Statements", "Update_date=#" & Format(Date, "mm/dd/yy") & "#") > 0 Then
        CurrentDb.Execute "Delete Tbl_Historical_Statements.Customer_ID FROM Tbl_Historical_Statements WHERE (([Tbl_Historical_Statements].[Update_date]=Date()));"
    End If

    PrevMonthEnd = Format(DMax("[MonthEnd]", "[tbl_MonthEnd]", "MonthEnd<#" & Format(Date, "mm/dd/yy") & "#"), "mm/dd/yy")

    Set rst = CurrentDb.OpenRecordset("SELECT Tbl_Historical_Statements.Update_date FROM Tbl_Historical_Statements GROUP BY Tbl_Historical_Statements.Update_date HAVING (((Tbl_Historical_Statements.Update_date)>#" & Format(PrevMonthEnd, "mm/dd/yy") & "#));")
    If rst.RecordCount < 2 Then
        CurrentDb.Execute "Insert INTO Tbl_Historical_Statements ( Customer_ID ,Update_date, [Date], Document_Number, Customer_reference, Type, Amount, Overdue_Date, [Currency]) SELECT Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date, Tbl_Invoices.Date, Tbl_Invoices.Date, Tbl_Invoices.Document_Number , Tbl_Invoices.Type, Tbl_Invoices.Amount, Tbl_Invoices.Overdue_Date, Tbl_Invoices.Currency  FROM Tbl_Invoices WHERE ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#);"
    Else
        For I = rst.RecordCount - 1 To 1 Step -1
            CurrentDb.Execute "Delete Tbl_Historical_Statements.Customer_ID FROM Tbl_Historical_Statements WHERE (([Tbl_Historical_Statements].[Update_date]=#" & Format(DMax("Update_date", "Tbl_Historical_Statements"), "mm/dd/yy") & "#));"
        Next I
        CurrentDb.Execute "Insert INTO Tbl_Historical_Statements ( Customer_ID ,Update_date, [Date], Document_Number, Customer_reference, Type, Amount, Overdue_Date, [Currency]) SELECT Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date, Tbl_Invoices.Date, Tbl_Invoices.Date, Tbl_Invoices.Document_Number , Tbl_Invoices.Type, Tbl_Invoices.Amount, Tbl_Invoices.Overdue_Date, Tbl_Invoices.Currency  FROM Tbl_Invoices WHERE ((Tbl_Invoices.Update_date)=#" & Format(Date, "mm/dd/yy") & "#);"
    End If
End Sub


Function EmailSelected(Position, Total) As Boolean
Dim a As Long
    If Position > Total Then
        EmailSelected = False
    Else
        EmailSelected = False
        While Total > 0
            a = DMax("[ID]", "[Tbl_EmailAddresses]", "ID<=" & Total)
            If a = Position Then
            If Position / a >= 1 Then
                EmailSelected = True
                Exit Function
            End If
            End If
            Total = Total - a
        Wend
    End If
End Function
Function FillEmailAddresses(Optional Code As Long) As String
Dim IDEmailAddress As Long
    FillEmailAddresses = ""
    If Code <> 0 Then
        While Code <> 0
            IDEmailAddress = DMax("[ID]", "[Tbl_EmailAddresses]", "ID<=" & Code)
            If Code / IDEmailAddress >= 1 Then
'                List61.AddItem (CStr(IDEmailAddress) & ";" & DLookup("[EmailAddress]", "[Tbl_EmailAddresses]", "ID=" & IDEmailAddress))
 '               FillEmailAddresses = FillEmailAddresses & CStr(IDEmailAddress) & ";" & DLookup("[EmailAddress]", "[Tbl_EmailAddresses]", "ID=" & IDEmailAddress) & ";"
                Code = Code - IDEmailAddress
            End If
        Wend
    End If
End Function




Sub AttachDocumentsToInvoices(OriginForm As Variant)
    Dim StrFilter As String
    Dim StrInputFileName As Variant
    Dim lngFlags As Long

    StrFilter = ahtAddFilterItem(StrFilter, "All files(*.*, *.*)", "*.*")
    lngFlags = ahtOFN_ALLOWMULTISELECT Or ahtOFN_EXPLORER
    StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
    DialogTitle:="Please select file to link", Flags:=lngFlags)
    If lngFlags And ahtOFN_ALLOWMULTISELECT Then
        On Error Resume Next
        If StrInputFileName <> "" Then
            If IsArray(StrInputFileName) Then 'multiple attach files selected
                Dim I As Integer
                For I = 0 To UBound(StrInputFileName)
                    Call AddNewInvoiceAttachment(StrInputFileName(I), OriginForm)
                Next I
            Else '1 file attach selected
               Call AddNewInvoiceAttachment(StrInputFileName, OriginForm)
            End If
        End If
        On Error GoTo 0
    End If
End Sub

Sub AddNewInvoiceAttachment(Attachment As Variant, OriginForm As Variant)
    Dim rst As Recordset
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = CurrentProject.Connection
        .Open "Tbl_InvoiceAttachments", , adOpenKeyset, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("AttachName") = DLookup("[PathInvoiceAttachments]", "TblGeneral") & _
            Left(Dir(Attachment), InStrRev(Dir(Attachment), ".") - 1) & " D" & _
            Format(Date, "ddmmyyyy") & " T" & Format(Time(), "hhmmss") & _
            Mid(Dir(Attachment), InStrRev(Dir(Attachment), "."), 100)
        .Fields("CustomerID") = OriginForm.Recordset("Customer_ID")
        '.Fields("DocumentID") = Left(OriginForm.Recordset("Document_Number") & "       ", 7) & _
            Left(OriginForm.Recordset("Date") & "          ", 10) & OriginForm.Recordset("Type") & _
            Left(OriginForm.Recordset("Customer_reference") & "                 ", 15)

        .Fields("DocumentID") = GetDocumentsToInvoices(OriginForm)
        .Update
        OriginForm.Recordset.Edit
        OriginForm.Recordset("Attachment") = True
        OriginForm.Recordset.Update
        FileCopy Attachment, .Fields("AttachName")
        .Close
    End With
End Sub

Sub ShowAttachDocumentsToInvoices(OriginForm As Variant)
    Dim rst As Variant
    Rem show attachments linked to invoices open in the statement
    Set rst = New ADODB.Recordset
    With rst
'        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_InvoiceAttachments.AttachName, Tbl_InvoiceAttachments.DocumentID, Tbl_InvoiceAttachments.CustomerID  FROM Tbl_InvoiceAttachments WHERE (((Tbl_InvoiceAttachments.CustomerID)=" & OriginForm.Recordset("Customer_ID") & ") AND ((Tbl_InvoiceAttachments.DocumentID)='" & _
            Left(OriginForm.Recordset("Document_Number") & "       ", 7) & _
            Left(OriginForm.Recordset("Date") & "          ", 10) & OriginForm.Recordset("Type") & _
            Left(OriginForm.Recordset("Customer_reference") & "                 ", 15) & "'));")
        Set rst = CurrentDb.OpenRecordset("SELECT Tbl_InvoiceAttachments.AttachName, Tbl_InvoiceAttachments.DocumentID, Tbl_InvoiceAttachments.CustomerID  FROM Tbl_InvoiceAttachments WHERE (((Tbl_InvoiceAttachments.CustomerID)=" & OriginForm.Recordset("Customer_ID") & ") AND ((Tbl_InvoiceAttachments.DocumentID)='" & GetDocumentsToInvoices(OriginForm) & "'));")

    End With
    While Not rst.EOF
        Application.FollowHyperlink rst.Fields("AttachName"), , True
        rst.MoveNext
    Wend
End Sub


Function GetDocumentsToInvoices(OriginForm As Variant) As String
    GetDocumentsToInvoices = _
        Left(OriginForm.Recordset("Document_Number") & "       ", 7) & _
        Left(OriginForm.Recordset("Date") & "          ", 10) & OriginForm.Recordset("Type") & _
        Left(OriginForm.Recordset("Customer_reference") & "                 ", 15)
End Function

Function ExcelSheetExists(SheetName As String, WB As Variant) As Boolean
    Dim wksht As Worksheet
    ExcelSheetExists = False
    For Each wksht In WB.Worksheets
        If wksht.Name = SheetName Then
            ExcelSheetExists = True
            Exit For
        End If
    Next wksht
End Function

Sub AddPaymentsReceived(ExcApp As Excel.Application, doc As Excel.Workbook, CustomerID)
Dim I As Integer
Dim rst As Variant
Dim a As String
With ExcApp
    Set rst = New ADODB.Recordset
    a = "SELECT Tbl_CashCollected.CustomerID, Tbl_CashCollected.[Payment Date], Tbl_CashCollected.Currency, Tbl_CashCollected.Amount, Tbl_CashCollected.[Original amount]  " & _
        " FROM Tbl_CashCollected WHERE (((Tbl_CashCollected.CustomerID)='" & CustomerID & "') AND ((Tbl_CashCollected.[Payment Date])>#" & Format(Date - 180, "mm/dd/yy") & "#)) ORDER BY Tbl_CashCollected.[Payment Date] DESC"
    Set rst = CurrentDb.OpenRecordset(a)

    If rst.RecordCount > 0 Then
        rst.MoveFirst
        I = 1
        While InStr(1, ExcApp.Sheets(I).Name, "Sheet") = 0
            I = I + 1
        Wend
        I = I - 1
        doc.Worksheets.Add(After:=doc.Worksheets(I)).Name = "Last 6 months payments "
        I = 1
        ExcApp.ActiveSheet.Cells(I, 1) = "Payment date"
        ExcApp.ActiveSheet.Cells(I, 2) = "Currency"
        ExcApp.ActiveSheet.Cells(I, 3) = "Amount"

        For I = 0 To rst.RecordCount - 1
            ExcApp.ActiveSheet.Cells(I + 2, 1) = Format(rst.Fields("Payment Date"), "dd-mmm-yyyy")
            ExcApp.ActiveSheet.Cells(I + 2, 2) = rst.Fields("Currency")
            ExcApp.ActiveSheet.Cells(I + 2, 3) = Format(rst.Fields("Original amount"), "##,##0.00")
            rst.MoveNext
        Next I
        ExcApp.ActiveSheet.Range("A..C").Columns.AutoFit
        ExcApp.Sheets(1).Select
    End If
End With
End Sub

Sub SyncReplica()
Dim dbSynch As DAO.Database
Dim strLocal As String
Dim strRemote As String
Dim timea, timeb As Date
    Rem Call CompattaBE
    timea = Now
    On Error GoTo Err_Command0_Click
    strLocal = "C:\Access\Replica of db2.mdb"
    strRemote = "Q:\Credit Control\Access\db2.mdb"
    Set dbSynch = DBEngine(0).OpenDatabase(strLocal)
    dbSynch.Synchronize strRemote, dbRepImpExpChanges
    dbSynch.Close
    Set dbSynch = Nothing
    MsgBox "Synch done in: " & Format(Now - timea, "hh:mm:ss")
Exit_Command0_Click:
    Exit Sub
Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Sub MakeCreditCheckFailureReport()
Dim ExcApp As Excel.Application
Dim ExcDoc As Excel.Workbook
Dim S As String
Dim Rec As Variant
Dim I, a As Integer
    I = 0
    Set ExcApp = CreateObject("Excel.Application") 'apre il modello di Excel
    Set ExcDoc = ExcApp.Workbooks.Open(GetPathExcelDirectory & "Failed Credit Check Releases Template.XLS")
    ExcApp.Visible = True
    With ExcDoc.Sheets(1)
        S = .Cells(4, 3)
        S = Replace(S, "TO BE REPLACED", Format(DLookup("[Update_CL+1]", "[tblgeneral]"), "dd-mmm-yyyy"))
        .Cells(4, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE ((Not (Tbl_credit_check_failures.[Hold Name])='LOGI Manual Credit Hold')); ")
        .Cells(4, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        S = .Cells(6, 3)
        S = Replace(S, "TO BE REPLACED", Format(DLookup("[Update_CL+1]", "[tblgeneral]"), "dd-mmm-yyyy"))
        .Cells(6, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE ((Not (Tbl_credit_check_failures.[Hold Name])='LOGI Manual Credit Hold') AND ((Tbl_credit_check_failures.Released)=True));")
        .Cells(6, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR, Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.Released, Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID GROUP BY Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.Released, Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country HAVING (((Tbl_credit_check_failures.[Hold Name]) <> 'LOGI Manual Credit Hold') And ((Tbl_credit_check_failures.Released) = False)) ORDER BY Sum([Amount]*[ExchangeRate]) DESC;")
        If Rec.RecordCount > 0 Then
            Rec.MoveFirst
            I = 0
            While Not Rec.EOF
                .Cells(12 + I, 3) = Rec.Fields("country")
                .Cells(12 + I, 4) = DLookup("Name", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                .Cells(12 + I, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")
                .Cells(12 + I, 6) = DLookup("ReleaseNotes", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                I = I + 1
                Rec.MoveNext
            Wend
            .Range("A" & 12 + I & ":z100").Delete Shift:=xlUp
            I = I + 14
        End If
        If I = 0 Then I = 103
        S = .Cells(I, 3)
        S = Replace(S, "TO BE REPLACED", Format(DLookup("[Update_CL+1]", "[tblgeneral]"), "dd-mmm-yyyy"))
        .Cells(I, 3) = S
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID WHERE (((Tbl_credit_check_failures.[Hold Type])='LOGI Manual Credit Hold'));")
        .Cells(I, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")

        I = I + 4
        Set Rec = CurrentDb.OpenRecordset("SELECT Sum([Amount]*[ExchangeRate]) AS AmountInEUR, Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country, Tbl_credit_check_failures.[Hold Type] FROM Tbl_credit_check_failures INNER JOIN Tbl_Currencies ON Tbl_credit_check_failures.[Currency Code] = Tbl_Currencies.CurrencyID GROUP BY Tbl_credit_check_failures.[Hold Name], Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Customer Number], Tbl_credit_check_failures.Country, Tbl_credit_check_failures.[Hold Type] HAVING (((Tbl_credit_check_failures.[Hold Type])='LOGI Manual Credit Hold')) ORDER BY Sum([Amount]*[ExchangeRate]) DESC;")
        If Rec.RecordCount > 0 Then
            Rec.MoveFirst
            a = 0
            While Not Rec.EOF
                .Cells(I + a, 3) = Rec.Fields("country")
                .Cells(I + a, 4) = DLookup("Name", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                .Cells(I + a, 5) = Format(Rec.Fields("AmountInEUR") / 1000, "##,##0.00")
                .Cells(I + a, 6) = DLookup("ReleaseNotes", "tbl_customers", "Customer_code=" & Rec.Fields("Customer Number"))
                I = I + 1
                Rec.MoveNext
            Wend
            .Range("A" & I & ":z300").Delete Shift:=xlUp
        End If
        .Range("A1").Select
    End With
    S = GetPathExcelDirectory & "Failed Credit Check Releases Today.xls"

    ExcApp.Application.DisplayAlerts = False
    ExcApp.ActiveWorkbook.SaveAs FileName:=S, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    ExcApp.Application.DisplayAlerts = True
    ExcApp.Quit
    Set ExcApp = Nothing
    Set ExcDoc = Nothing

End Sub

Public Function ExportTextDelimited(strQueryName As String, strDelimiter As String)

Dim RS          As Variant
Dim strHead     As String
Dim strData     As String
Dim inti        As Integer
Dim intFile     As Integer
Dim fso         As New FileSystemObject

On Error GoTo Handle_Err

    fso.CreateTextFile ("C:\" & strQueryName & ".csv")

    Set RS = CurrentDb.OpenRecordset(strQueryName)

    RS.MoveFirst

    intFile = FreeFile
    strHead = ""

    'Add the Headers
    For inti = 0 To RS.Fields.Count - 1
        If strHead = "" Then
            strHead = RS.Fields(inti).Name
        Else
            strHead = strHead & strDelimiter & RS.Fields(inti).Name
        End If
    Next

    Open "C:\" & strQueryName & ".csv" For Output As #intFile

    Print #intFile, strHead

    strHead = ""

    'Add the Data
    While Not RS.EOF

        For inti = 0 To RS.Fields.Count - 1
            If strData = "" Then
                Rem strData = IIf(IsNumeric(rs.Fields(inti).value), rs.Fields(inti).value, IIf(IsDate(rs.Fields(inti).value), rs.Fields(inti).value, """" & rs.Fields(inti).value & """"))
                strData = IIf(IsNumeric(RS.Fields(inti).value), RS.Fields(inti).value, IIf(IsDate(RS.Fields(inti).value), RS.Fields(inti).value, RS.Fields(inti).value))
            Else
                Rem strData = strData & strDelimiter & IIf(IsNumeric(rs.Fields(inti).value), rs.Fields(inti).value, IIf(IsDate(rs.Fields(inti).value), rs.Fields(inti).value, """" & rs.Fields(inti).value & """"))
                strData = strData & strDelimiter & IIf(IsNumeric(RS.Fields(inti).value), RS.Fields(inti).value, IIf(IsDate(RS.Fields(inti).value), RS.Fields(inti).value, RS.Fields(inti).value))
            End If
        Next

        Print #intFile, strData

        strData = ""

        RS.MoveNext
    Wend

        Close #intFile

RS.Close
Set RS = Nothing

'Open the file for viewing
Rem Application.FollowHyperlink "C:\" & strQueryName & ".csv"

Exit Function

Handle_Err:
MsgBox Err & " - " & Err.Description
End Function
