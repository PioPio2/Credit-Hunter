Attribute VB_Name = "Utility2"
Option Compare Database

Rem *** There can be only one  !!! ***

Rem note rel. ######## 1.1.9 ########
Rem Added to send "offender letter" option in the who paid yesterday

Rem note rel. ######## 1.2 ########
Rem BUG FIXED: notes in Order release procedure are sorted now
Rem BUG FIXED: wrong customer name in first email after import fixed
Rem BUG FIXED: Excel statement file fixed (the column G was too narrow sometimes)
Rem BUG FIXED: duplicate chargeback lines fixed (in chargeback file sometimes it adds the chargeback more than once
Rem added FindCustomerLastDate in Scheduler Onopen, it should fix the issue of the exposure=0
Rem added field "LatestPaymentDate" in "Tbl_MonthEnd" table and added ID n# 15 in DivideTemplateInBits procedure in order to change the latest payment day in the forward aging
Rem BUG FIXED: duplicate Chargebacks in Excel file

Rem note rel. ######## 1.2.5 ########
Rem BUG FIXED: duplicate invoices in history fixed
Rem offender letter option inserted in "who paid yesterday" and "who paid yesterday and even before"
Rem DSO calculation inserted

Rem note rel. ######## 1.2.7 ########
Rem BUG FIXED: Comments in who paid yesterday are sorted by date now
Rem BUG FIXED: DSO calculation fixed
Rem added DSO report on annual basis
Rem BUG FIXED: from now on every time you import new CL from Atradius the old ones will be erased
Rem added insurance CL in scheduler

Rem note rel. ######## 1.2.7.7 ########
Rem BUG FIXED: Chargeback issue fixed
Rem Master Data File Button inserted in Scheduler
Rem BUG FIXED: Fixed error on statement total amount >31-60 days
Rem Added tab "master data file" in the scheduler

Rem note rel. ######## 1.3 ########
Rem "Open Master Data file" button moved from "Search" to "Customer header" tab.
Rem Box "status" in scheduler main form  is only for display from now on.
Rem "Customer header" tab has been created with two new information: DSO and Static notes (see red ellipses).
Rem template for each customer status. Automatic delivery of emails on changing status

Rem note rel. ######## 1.4.1.2 ########
Rem now if you send the forward aging the cc field includes both internal and external contact. Previously only external contacts were included
Rem BUG FIXED:  intercompany invoices are deleted from the archive
Rem BUG FIXED: Now the Account overview shows ALL customers with overdue > 30 days and not only the ones with status <> nul
Rem now the who paid yesterday and who paid yesterday and even before have both the status of the customer
Rem BUG FIXED: Now if you select a new status and move only the tab the email text doesn't change anymore
Rem BUG FIXED: Now all the cc people in the who paid yesterday and even before + forward statement receive the email
Rem From now on the release horizon overview will be 3 days. Changed release procedure + scheduler accordingly

Rem note rel. ######## 1.4.1.5 ########
Rem BUG FIXED: From now on the Area will be included in the Tbl_Customer table
Rem BUG FIXED: From now on the data in Tbl_AdditionalQueryData will be inserted correctly
Rem MakeGeneralQueryLogFile Added (new report for query log)

Rem note rel. ######## 1.4.1.7 ########
Rem CL Report in Excel format is available now. (option to send automatically by email)
Rem General Query Log file moved to Utility menu
Rem From now on everytime the CL is uploaded there's an additional control on the Thresold date. If it's different than today+3 days there's a warning
Rem Fixed the Scheduler issue following the installation of Ms Office 2007
Rem Bug fixed the write off report is including also the customers without any procedure open now.
Rem Bug fixed. The order release request email is sent to the correct person now. Before the procedure was looking at the excess amount in USD currency instead of EUR
Rem Bug fixed. Access sends the release requests to the right people.

Rem note rel. ######## 1.4.2.1 ########
Rem Amended CL horizon, from +5 to +7 days

Rem note rel. ######## 1.4.2.3 ########
Rem Bug fixed. Now all customers in the CL report appears in the CL report Excel format

Rem note rel. ######## 1.4.2.4 ########
Rem Make up

Rem note rel. ######## 1.4.2.5 ########
Rem Bug fixed. From now on writeoffs report shows either documents with date<next month end -365 days or documents with date >next month end-365 BUT with write offs in notes


Rem note rel. ######## 1.4.2.6 ########
Rem Bug fixed. From now The release requests go to the right person

Rem note rel. ######## 1.4.3.0 ########
Rem removed ShowUserRosterMultipleUsers

Rem note rel. ######## 1.4.2.7 ########
Rem Bug fixed. From now The query log contains also the queries without any date

Rem note rel. ######## 1.4.2.9 ########
Rem Bug fixed. Approval email addresses are ok now

Rem rel. ########### 1.8.0.0 ############
Rem Function NumMaxRows improved (dichotomic search)
Rem Chargeck statement uploading (chargeback side) improved
Rem Bug fixed. Query log data insert+chargebacks in Excel file from scheduler fixed
Rem Releases+Statement files are saved in "Document" folder now
Rem query log file has no more chargebacks and credit notes
Rem added $mart import information procedure
Rem From now on customer with no statement have next appointment date = tomorrow

Rem rel. ########### 1.8.0.1 ############
Rem moved again the releases orizon to 5 days

Rem rel. ########### 1.8.0.2 ############
Rem Write off procedure looks at invoices older than 180 days

Rem rel. ########### 1.9.0.0 ############
Rem introduced superuser field
Rem introduced dashboard
Rem Inserted Target information

Rem rel. ########### 1.9.0.1 ############
Rem fixed the problem with the approver names/email addresses
Rem added chart of the disputes in the dashboard
Rem if you double click on the chart in the dashboard then you have more details of it
Rem added the field "to be released" and the report is printed

Rem rel. ########### 1.9.27 ############
Rem new feature. Option to select which query+comment to be printed in the staement
Rem new feature. Option to Attach files to the invoices in teh scheduler

Rem rel. ########### 1.9.31 ############
Rem new feature. From now on the InsuredCL will be uploaded from the Collection Management report

Rem rel. ########### 1.9.32 ############
Rem new feature. Removed three options in the scheduler concerning the statement run. They were useless.
Rem improvement. From now on you can't send out release requests if Leilani's file is uploaded the day after the Credit limit report

Rem rel. ########### 1.9.40 ############
Rem Bug Fixed. When all customers next appointments were set in the future there was no access to the scheduler. Fixed
Rem Improvement. From now on you can upload either Leilani's file or the LOGI_Hold_Report_New from Oracle from the same menu.

Rem rel. ########### 1.9.42 ############ Aug 8th 2011
Rem added email GMail setup
Rem installed procedure in Fremont offices (AMR region)
Rem improvement. amended the on account report. Now it shows data even if the format date is english mm/dd/yy

Rem rel. ########### 1.9.50 ############
Rem Improvement. The scheduler has now the filter and the sort by columns option. This is reflected in the statements
Rem improvement. The currency is managed according to the user's pc Windows Setup
Rem Improvement. You can set the main currency to be used thorough the application. Before it was EUR only
Rem bug fixed: Total overdue amount on month end in the scheduler is working all the time now
Rem improvement: dashboard-pie chart. Replaced label legal with third part collection
Rem improvement: amended release excel form from CL excess with CL override
Rem improvement: the number of the statements in the DB (useful in the who paid yesterday and even before) can be changed
Rem improvement: the aging in the scheduler and the CLs information have been swapped in the scheduler. Confusing at the beginning but the aging is aligned to the invoice values
Rem improvement: option to include the user's email address in the recipients of all messages sent out by Access DB. To be used when messages are not saved in the GMail sent items automatically.
Rem improvement: import icons have the report name for best understanding

Rem rel. ########### 1.9.77 ############
Rem improvement: Cash collected import procedure re-written to improve statemnet upload performance. From now on thewhopaidyesterday is according tothe cash collected date
Rem bug fixed: The who paid yesterday didn't work properly before
Rem bug fixed: Who paid yesterday & even before and forward aging are back working now
Rem bug fixed: The aging in the scheduler is working now

Rem rel. ########### 1.9.90 ############
Rem improvement: One can upload more than one LOGI_Hold_Report_New now
Rem improvement: More than one statement/CL report/customer failing the credit check report (only .txt) /cash collected report can be imported at the same time
Rem improvement: Wording message adjusted after upload
Rem improvement: added Main Currency in General table
Rem improvement: added Area in General table to identify EMEA AMR etc

Rem improvement: added table Channel in DB for report purposes
Rem bug fixed: Invoice customer reference with letters non included in ANSI set å,ö,ä,ü etc. are now imported correctly during the statement upload
Rem improvement: The import process is much faster now

Rem rel. ########### 1.9.91 ############
Rem bug fixed: the merging file routine is working properly now

Rem rel. ########### 1.9.97 ############
Rem local time not maintained any more. Replaced with Sales Channel
Rem improvement: new cash target report layout upload
Rem improvement: added extra field email target reports in General table
Rem bug fixed:as not all payment downloads refer to the previous day an am endment was necessary to make the who paid yesterday working properly
Rem bug fixed: the automatic email that highlights the disappeared payments has been fixed now.
Rem improvement: new ESDC statement layout upload
Rem new statement now with additional information: so number, factura number, pull ticket number, original invoice amount.
    Rem of course the scheduler is different.
    '- So number appears all the time in both, scheduler and statement
    '- Factura number and pull ticket number appear in the scheduler only if there is data in it and in the statement only if the flag is set
    '- Original invoice amount. It appears all the time in thescheduler and in the statement only if the flag is set
Rem in the setup created 3 boolean fields (factura number, pull ticket number and original invoice amount) that decide if to print the info or not in the statement

Rem rel. ########### xxxxx ############
Rem improgement: the cash collected report email has the picture of the attachment in the text+HTML format
Rem bug fixed: in the who paid yesterday the total amount of the payment appears now
Rem improvement: added column currency and original amount and one more pivot table in cash collected report
Rem improvement: replaced the word "terminated with "completed" in the import messages


Private Type STARTUPINFO
cb As Long
lpReserved As String
lpDesktop As String
lpTitle As String
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type

Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessID As Long
dwThreadID As Long
End Type

Rem Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
lpStartupInfo As STARTUPINFO, lpProcessInformation As _
PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    StrFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    StrFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Stringa As String

Const LUNGHEZZA_MASSIMA_PERCORSO = 255

Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare PtrSafe Function aht_apiGetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean

Declare PtrSafe Function aht_apiGetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
' You won't use these.
'Global Const ahtOFN_ENABLEHOOK = &H20
'Global Const ahtOFN_ENABLETEMPLATE = &H40
'Global Const ahtOFN_ENABLETEMPLATEHANDLE = &H80
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
' New for Windows 95
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Function TestIt()
    Dim StrFilter As String
    Dim lngFlags As Long
    StrFilter = ahtAddFilterItem(StrFilter, "Access Files (*.mda, *.mdb)", _
                    "*.MDA;*.MDB")
    StrFilter = ahtAddFilterItem(StrFilter, "dBASE Files (*.dbf)", "*.DBF")
    StrFilter = ahtAddFilterItem(StrFilter, "Text Files (*.txt)", "*.TXT")
    StrFilter = ahtAddFilterItem(StrFilter, "All Files (*.*)", "*.*")
    MsgBox "You selected: " & ahtCommonFileOpenSave(InitialDir:="C:\", _
        Filter:=StrFilter, FilterIndex:=3, Flags:=lngFlags, _
        DialogTitle:="Hello! Open Me!")
    ' Since you passed in a variable for lngFlags,
    ' the function places the output flags value in the variable.
    Debug.Print Hex(lngFlags)
End Function

Function GetOpenFile(Optional varDirectory As Variant, _
    Optional varTitleForDialog As Variant) As Variant
' Here's an example that gets an Access database name.
Dim StrFilter As String
Dim lngFlags As Long
Dim varFileName As Variant
' Specify that the chosen file must already exist,
' don't change directories when you're done
' Also, don't bother displaying
' the read-only box. It'll only confuse people.
    lngFlags = ahtOFN_FILEMUSTEXIST Or _
                ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
    If IsMissing(varDirectory) Then
        varDirectory = ""
    End If
    If IsMissing(varTitleForDialog) Then
        varTitleForDialog = ""
    End If

    ' Define the filter string and allocate space in the "c"
    ' string Duplicate this line with changes as necessary for
    ' more file templates.
    StrFilter = ahtAddFilterItem(StrFilter, _
                "Access (*.mdb)", "*.MDB;*.MDA")
    ' Now actually call to get the file name.
    varFileName = ahtCommonFileOpenSave( _
                    OpenFile:=True, _
                    InitialDir:=varDirectory, _
                    Filter:=StrFilter, _
                    Flags:=lngFlags, _
                    DialogTitle:=varTitleForDialog)
    If Not IsNull(varFileName) Then
        varFileName = TrimNull(varFileName)
    End If
    GetOpenFile = varFileName
End Function

Function ahtCommonFileOpenSave( _
            Optional ByRef Flags As Variant, _
            Optional ByVal InitialDir As Variant, _
            Optional ByVal Filter As Variant, _
            Optional ByVal FilterIndex As Variant, _
            Optional ByVal DefaultExt As Variant, _
            Optional ByVal Filename As Variant, _
            Optional ByVal DialogTitle As Variant, _
            Optional ByVal hwnd As Variant, _
            Optional ByVal OpenFile As Variant) As Variant
' This is the entry point you'll use to call the common
' file open/save dialog. The parameters are listed
' below, and all are optional.
'
' In:
' Flags: one or more of the ahtOFN_* constants, OR'd together.
' InitialDir: the directory in which to first look
' Filter: a set of file filters, set up by calling
' AddFilterItem. See examples.
' FilterIndex: 1-based integer indicating which filter
' set to use, by default (1 if unspecified)
' DefaultExt: Extension to use if the user doesn't enter one.
' Only useful on file saves.
' FileName: Default value for the file name text box.
' DialogTitle: Title for the dialog.
' hWnd: parent window handle
' OpenFile: Boolean(True=Open File/False=Save As)
' Out:
' Return Value: Either Null or the selected filename
Dim OFN As tagOPENFILENAME
Dim strfilename As String
Dim strFileTitle As String
Dim fResult As Boolean
    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(Filename) Then Filename = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strfilename = Left(Filename & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .StrFilter = Filter
        .nFilterIndex = FilterIndex
        .StrFile = strfilename
        .nMaxFile = Len(strfilename)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If

 ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.
    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        If Flags And ahtOFN_ALLOWMULTISELECT Then
            ' Return the full array.
            Dim items As Variant
            Dim value As String
            value = OFN.StrFile
            ' Get rid of empty items:
            Dim I As Integer
            For I = Len(value) To 1 Step -1
              If Mid$(value, I, 1) <> Chr$(0) Then
                Exit For
              End If
            Next I
            value = Mid(value, 1, I)

            ' Break the list up at null characters:
            items = Split(value, Chr(0))

            ' Loop through the items in the "array",
            ' and build full file names:
            Dim numItems As Integer
            Dim Result() As String

            numItems = UBound(items) + 1
            If numItems > 1 Then
                ReDim Result(0 To numItems - 2)
                For I = 1 To numItems - 1
                    Result(I - 1) = FixPath(items(0)) & items(I)
                Next I
                ahtCommonFileOpenSave = Result
            Else
                ' If you only select a single item,
                ' Windows just places it in item 0.
                ahtCommonFileOpenSave = items(0)
            End If
        Else
            ahtCommonFileOpenSave = TrimNull(OFN.StrFile)
        End If
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function


Function ahtAddFilterItem(StrFilter As String, _
    strDescription As String, Optional varItem As Variant) As String
' Tack a new chunk onto the file filter.
' That is, take the old value, stick onto it the description,
' (like "Databases"), a null character, the skeleton
' (like "*.mdb;*.mda") and a final null character.

    If IsMissing(varItem) Then varItem = "*.*"
    ahtAddFilterItem = StrFilter & _
                strDescription & vbNullChar & _
                varItem & vbNullChar
End Function

Private Function TrimNull(ByVal strItem As String) As String
Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = Left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function

Private Function FixPath(ByVal Path As String) As String
    If Right$(Path, 1) <> "\" Then
        FixPath = Path & "\"
    Else
        FixPath = Path
    End If
End Function


Public Function NomeFileDos(szPercorso As String) As String
    Dim lLunghezza As Long
    NomeFileDos = String(LUNGHEZZA_MASSIMA_PERCORSO, vbNullChar)
    lLunghezza = GetShortPathName(szPercorso, NomeFileDos, Len(szPercorso))
    NomeFileDos = Left$(NomeFileDos, lLunghezza)
    If Asc(NomeFileDos) = 0 Then
        NomeFileDos = szPercorso
    End If
End Function
