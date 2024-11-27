Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Support As IWhoPaidYesterday
Private SkipOnCurrent As Boolean
Private Sub SendSingleStatement()
    On Error GoTo restoremouse
    DoCmd.Hourglass True
    Dim StatementSupport As iStatementSupport
    Set StatementSupport = New clsStatementSupportLIVE    'create connection with DB

    Dim CustomerHeader As clsCustomerHeader
    Set CustomerHeader = New clsCustomerHeader    ' collect and populate customer master data
    Call CustomerHeader.Populate(Me.Recordset.Fields("Customer_code").value)

    If Not IsNull(CustomerHeader.getEmail) Then    ' if the master data has an email address then statement and send to customer
        Dim Statement As IStatement
        Set Statement = New clsStatement
        Dim Overdue As Currency
        Overdue = 0
        Dim Outstanding As Currency
        Outstanding = 0
        Dim outCurrency As String
        Call Statement.CreateStatement(StatementSupport, CustomerHeader, Outstanding, Overdue, outCurrency, True)

        Dim StatementPath As String
        StatementPath = Statement.GetPathStatement    ' get full file name where the updated statement is saved

        Dim OL As IOutlook
        Set OL = New clsOutlook
        OL.CreateOutlook

        Dim Details As clsCustomerDetails
        Set Details = New clsCustomerDetails
        Call Details.Populate(Outstanding, Overdue, outCurrency)

        Dim Attachments As New Collection
        Attachments.Add (Statement.GetPathStatement)

        Dim WordPath As String
        WordPath = Application.CurrentProject.path & "\Templates\Word\" & "Statement.docx"

        Dim Subject As String
        Subject = "Updated statement " & CustomerHeader.getCustomerName
        'Dim Result As Boolean
        Call OL.SendEmailFromTemplate(Subject, CustomerHeader, Details, Attachments, WordPath, False)
    Else
        MsgBox ("No email address for this customer")
    End If

    DoCmd.Hourglass False
    Set StatementSupport = Nothing
    Set CustomerHeader = Nothing
    Set Statement = Nothing
    Set Details = Nothing
    Exit Sub

restoremouse:
    DoCmd.Hourglass False
    Set StatementSupport = Nothing
    Set CustomerHeader = Nothing
    Set Statement = Nothing
    Set Details = Nothing
End Sub

Private Sub btnSendAllStatements_Click()
    DoCmd.Hourglass True
    On Error GoTo ErrorHandler
    Me.Recordset.MoveFirst
    While Not Me.Recordset.EOF
        If Me.TabCtl48.Pages("Sheet1").Controls("Maschera1").Form.StatementExists Then
            Call SendSingleStatement
        End If
        Me.Recordset.MoveNext
    Wend
    DoCmd.Hourglass False
    MsgBox "All statements have been sent out", vbOKOnly
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
End Sub

Private Sub Button1_Click()
    Call SendSingleStatement
End Sub

Private Sub Form_Activate()
    DoCmd.Maximize
End Sub
Private Sub Form_Current()
    If SkipOnCurrent = True Then
        Me.Button1.enabled = Me.TabCtl48.Pages("Sheet1").Controls("Maschera1").Form.StatementExists
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    SkipOnCurrent = False
    Set Support = New clsWhoPaidYesterday
    Dim Injection As IWhoPaidYesterdaySupport
    Set Injection = New clsWhoPaidYesterdaySupportLIVE
    Call Support.Inject(Injection)

    Dim SQLString As String
    Call Support.UpdateSecondaryQueries          ' update two queries that retrieve the customer who paid something from the last import

    SQLString = Support.GetRecordSourceSQL
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQLString)
    If RS.RecordCount = 0 Then                   ' if nobody paid from the last report then exit Sub
        Call MsgBox("No customers paid yesterday", vbOKOnly)
        DoCmd.Close acForm, Me.Name, acSaveNo
        SkipOnCurrent = True
        Exit Sub
    End If

    ' otherwise keep on going and open the form to send statements to customer who paid the day before.
    Application.Echo False
    On Error Resume Next
    Me.RecordSource = SQLString
    Application.Echo True
    On Error GoTo 0
End Sub
