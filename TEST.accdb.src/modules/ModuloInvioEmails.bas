Attribute VB_Name = "ModuloInvioEmails"
Option Compare Database
Option Explicit

Type MapiMessage
    Reserved As Long
    Subject As String
    NoteText As String
    MessageType As String
    DateReceived As String
    ConversationID As String
    Flags As Long
    RecipCount As Long
    FileCount As Long
End Type

Type MapiRecip
    Reserved As Long
    RecipClass As Long
    Name As String
    Address As String
    EIDSize As Long
    EntryID As String
End Type

Type MapiFile
    Reserved As Long
    Flags As Long
    Position As Long
    PathName As String
    Filename As String
    FileType As String
End Type

Public Declare PtrSafe Function MAPISendMail Lib "MAPI32.DLL" _
        Alias "BMAPISendMail" _
        (ByVal Session&, _
        ByVal UIParam&, _
        Message As MapiMessage, _
        Recipient() As MapiRecip, _
        file() As MapiFile, _
        ByVal Flags&, _
        ByVal Reserved&) As Long
Public Declare PtrSafe Function GetProfileString Lib "kernel32" _
        Alias "GetProfileStringA" _
        (ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long) As Long
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

Global Const SUCCESS_SUCCESS = 0
Global Const MAPI_TO = 1
Global Const MAPI_CC = 2
Global Const MAPI_LOGON_UI = &H1

Public Function GetDefMail$()
    'verifico se esiste il supporto MAPI
    On Error GoTo Errore
    Dim Def$
    Dim di&

    Def$ = String$(3, 0)
    di = Nz(GetProfileString("Mail", "mapi", "", Def$, 2), "")
    'al posto di mapi si può ricercare "CMCDLLNAME32"
    'di = GetProfileString("Mail", "CMCDLLNAME32", "", Def$, 127)
    '---------non valido--------------------------Def$ = agGetStringFromLPSTR$(Def$)
    GetDefMail$ = di 'Def$
    Exit Function

Errore:
    MsgBox "Errore n° " & Err.Number & " - " & Err.Description
    GetDefMail$ = di
End Function

Public Function Mail()
    If GetDefMail$ = "" Or GetDefMail$ = "0" Then
        MsgBox "Non esiste un supporto MAPI per l'invio di posta elettronica"
        Exit Function
    End If
    Dim F As Form, Result
    Set F = Forms!Esporta
    If IsNull(F!To) Or F!To = "" Then Exit Function
    If IsNull(F!Subject) Then F!Subject = ""
    If IsNull(F!CC) Then F!CC = ""
    If IsNull(F!Attach) Then F!Attach = ""
    If IsNull(F!Message) Then F!Message = ""
    Result = SendMail((F!Subject), (F!To), (F!CC), (F!Attach), (F!Message))
    If Result <> SUCCESS_SUCCESS Then
        MsgBox "Errore nell'invio: " & Result, 16, "Mail"
    Else
        MsgBox "Operazione conclusa!", 64, "Mail"
    End If
End Function

Public Function SendMail(sSubject As String, sTo As String, sCC As String, sAttach As String, sMessage As String)
    Dim I, cTo, cCC, cAttach
    Dim MAPI_Message As MapiMessage

    cTo = CountTokens(sTo, ";")
    cCC = CountTokens(sCC, ";")
    cAttach = CountTokens(sAttach, ";")

    ReDim rTo(0 To cTo) As String
    ReDim rCC(0 To cCC) As String
    ReDim rAttach(0 To cAttach) As String

    ParseTokens rTo(), sTo, ";"
    ParseTokens rCC(), sCC, ";"
    ParseTokens rAttach(), sAttach, ";"

    ReDim MAPI_Recip(0 To cTo + cCC - 1) As MapiRecip

    For I = 0 To cTo - 1
        MAPI_Recip(I).Name = rTo(I)
        MAPI_Recip(I).RecipClass = MAPI_TO
    Next I

    For I = 0 To cCC - 1
        MAPI_Recip(cTo + I).Name = rCC(I)
        MAPI_Recip(cTo + I).RecipClass = MAPI_CC
    Next I

    ReDim MAPI_File(0 To cAttach) As MapiFile
    MAPI_Message.FileCount = cAttach

    For I = 0 To cAttach - 1
        MAPI_File(I).Position = -1
        MAPI_File(I).PathName = rAttach(I)
    Next I

    MAPI_Message.Subject = sSubject
    MAPI_Message.NoteText = sMessage
    MAPI_Message.RecipCount = cTo + cCC

    SendMail = MAPISendMail(0&, 0&, MAPI_Message, MAPI_Recip, MAPI_File, MAPI_LOGON_UI, 0&)
End Function

Public Function CountTokens(ByVal sSource As String, ByVal sDelim As String)
    Dim iDelimPos As Integer
    Dim iCount As Integer

    If sSource = "" Then
        CountTokens = 0
    Else
        iDelimPos = InStr(1, sSource, sDelim)
        Do Until iDelimPos = 0
            iCount = iCount + 1
            iDelimPos = InStr(iDelimPos + 1, sSource, sDelim)
        Loop
        CountTokens = iCount + 1
    End If
End Function

Public Function GetToken(sSource As String, ByVal sDelim As String) As String
    Dim iDelimPos As Integer
    iDelimPos = InStr(1, sSource, sDelim)
    If (iDelimPos = 0) Then
        GetToken = Trim$(sSource)
        sSource = ""
    Else
        GetToken = Trim$(Left$(sSource, iDelimPos - 1))
        sSource = Mid$(sSource, iDelimPos + 1)
    End If
End Function

Public Sub ParseTokens(a() As String, ByVal sTokens As String, ByVal sDelim As String)
    Dim I As Integer
    For I = LBound(a) To UBound(a)
        a(I) = GetToken(sTokens, sDelim)
    Next
End Sub



Function SendNotesMail(strTo As String, strcc As String, strSubject As String, strBody As String, strfilename As String, ParamArray strFiles())
    Dim doc As Object   'Lotus NOtes Document
    Dim rtitem As Object '
    Dim Body2 As Object
    Dim ws As Object    'Lotus Notes Workspace
    Dim oSess As Object 'Lotus Notes Session
    Dim oDB As Object   'Lotus Notes Database
    Dim x As Integer    'Counter
    'use on error resume next so that the user never will get an error
    'only the dialog "You have new mail" Lotus Notes can stop this macro
Do While fIsAppRunning = False
    MsgBox "Lotus Notes is not running" & Chr$(10) & "Make sure Lotus Notes is running and press OK."
Loop

On Error Resume Next

    Set oSess = CreateObject("Notes.NotesSession")
    'access the logged on users mailbox
    Set oDB = oSess.GETDATABASE("", "")
    Call oDB.OPENMAIL

    'create a new document as add text
    Set doc = oDB.CREATEDOCUMENT
    Set rtitem = doc.CREATERICHTEXTITEM("Body")
    doc.sendto = strTo
    doc.copyto = strcc
    doc.Subject = strSubject
    doc.body = strBody & vbCrLf & vbCrLf
    doc.SAVEMESSAGEONSEND = True

    'attach files
    If strfilename <> "" Then
        Set Body2 = rtitem.EMBEDOBJECT(1454, "", strfilename)
        If UBound(strFiles) > -1 Then
            For x = 0 To UBound(strFiles)
                Set Body2 = rtitem.EMBEDOBJECT(1454, "", strFiles(x))
            Next x
        End If
    End If
    doc.Send False
End Function

Sub Test()
Dim strTo, strcc As String        'The sendee(s) Needs to be fully qualified address. Other names seperated by commas
Dim strSubject As String    'The subject of the mail. Can be "" if no subject needed
Dim strBody As String       'The main body text of the message. Use "" if no text is to be included.
Dim FirstFile As String     'If you are embedding files then this is  the first one. Use "" if no files are to be sent
Dim SecondFile As String    'Add as many extra files as is needed, seperated by commas.
Dim ThirdFile As String     'And so on.

strTo = "acheriom@libero.it"
strcc = "alberto_paganini@libero.it"
strSubject = "Test Message"
strBody = "This is a test"
strBody = strBody & vbCrLf & "Just add new lines by concatenating vbCrLf "
FirstFile = "c:\A.TXT"
Rem SecondFile = "G:\life.xls"
Rem ThirdFile = "G:\CompactDbs.vbs"

'SendNotesMail strTo, strSubject, strBody, FirstFile
Rem , SecondFile, ThirdFile
End Sub

Private Function fIsAppRunning() As Boolean
'Looks to see if Lotus Notes is open
'Adapted from code by Dev Ashish

    Dim lngH As Long
    Dim lngX As Long, lngTmp As Long
    Const WM_USER = 1024
    On Local Error GoTo fIsAppRunning_Err
    fIsAppRunning = False

        lngH = apiFindWindow("NOTES", vbNullString)

    If lngH <> 0 Then
        apiSendMessage lngH, WM_USER + 18, 0, 0
        lngX = apiIsIconic(lngH)
        If lngX <> 0 Then
            lngTmp = apiShowWindow(lngH, 1)
        End If
        fIsAppRunning = True
    End If
fIsAppRunning_Exit:
    Exit Function
fIsAppRunning_Err:
    fIsAppRunning = False
    Resume fIsAppRunning_Exit
End Function
Sub InviaEmail(strTo As String, strcc As String, strSubject As String, strBody As String, strfilename As String, ParamArray strFiles())
    If fIsAppRunning = False Then
    Else
        SendEmails strTo, strcc, strSubject, strBody, strfilename, strFiles
    End If
End Sub

Sub SendNotesMail2(strTo As String, strcc As String, strSubject As String, strBody As String, strfilename As String, ParamArray strFiles())
    Dim doc As Object   'Lotus NOtes Document
    Dim rtitem As Object '
    Dim Body2 As Object
    Dim ws As Object    'Lotus Notes Workspace
    Dim oSess As Object 'Lotus Notes Session
    Dim oDB As Object   'Lotus Notes Database
    Dim x As Integer    'Counter
    Dim Recip() As String
    'use on error resume next so that the user never will get an error
    'only the dialog "You have new mail" Lotus Notes can stop this macro
Do While fIsAppRunning = False
    MsgBox "Lotus Notes is not running" & Chr$(10) & "Make sure Lotus Notes is running and press OK."
Loop

On Error Resume Next

    Set oSess = CreateObject("Notes.NotesSession")
    'access the logged on users mailbox
    Set oDB = oSess.GETDATABASE("", "")
    Call oDB.OPENMAIL

    'create a new document as add text
    Set doc = oDB.CREATEDOCUMENT
    Set rtitem = doc.CREATERICHTEXTITEM("Body")

    ReDim Recip(0 To CountTokens(strTo, ",")) As String
    ParseTokens Recip(), strTo, ","
    doc.sendto = Recip

    ReDim Recip(0 To CountTokens(strcc, ",")) As String
    ParseTokens Recip(), strcc, ","
    doc.copyto = Recip
    doc.Subject = strSubject
    doc.body = strBody & vbCrLf & vbCrLf
    doc.SAVEMESSAGEONSEND = True

    'attach files
    If strfilename <> "" Then
        Set Body2 = rtitem.EMBEDOBJECT(1454, "", strfilename)
        If UBound(strFiles) > -1 Then
            For x = 0 To UBound(strFiles)
                Set Body2 = rtitem.EMBEDOBJECT(1454, "", strFiles(x))
            Next x
        End If
    End If
    doc.Send False
End Sub

Sub InviaEmail2(strTo As String, strcc As String, strSubject As String, strBody As String, strfilename As String, ParamArray strFiles())
    If fIsAppRunning = False Then
    Else
'        SendEmails2 strTo, strcc, strSubject, strBody, strfilename
    End If
End Sub
Sub SendEmails(strTo As String, strcc As String, strSubject As String, strBody As String, strfilename As String, ParamArray strFiles())
Dim cdomsg As Variant
Dim Rst As Recordset
Dim x As Integer
If strTo <> "" Then
    Set Rst = New ADODB.Recordset
    Rst.ActiveConnection = CurrentProject.Connection
    Rst.Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set cdomsg = CreateObject("CDO.message")
    With cdomsg.Configuration.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = Rst.Fields("sendusing") 'NTLM method
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Rst.Fields("smtpserver")
        .item("http://schemas.microsoft.com/cdo/configuration/smptserverport") = Rst.Fields("SMTPserverport")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Rst.Fields("smtpauthenticate")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Rst.Fields("smtpusessl")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Rst.Fields("smtpconnectiontimeout")
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = DLookup("[Password]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .Update
    End With
    ' build email parts
    With cdomsg
        .To = strTo
        '.to = "alberto_paganini@libero.it"
        If strcc <> "" Then
            .CC = strcc
        End If
       ' .cc = ""
        If DLookup("[EmailSentToSender]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'") Then
            .CC = .CC & "," & DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        End If
        .From = DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .Subject = strSubject
        .TextBody = strBody

        'add attachments
        If strfilename <> "" Then
            .Addattachment strfilename
            If UBound(strFiles) > -1 Then
                For x = 0 To UBound(strFiles)
                    .Addattachment strFiles(x)
                Next x
            End If
        End If
        .Send
    End With
    Rst.Close
    Set Rst = Nothing
    Set cdomsg = Nothing
Else
    x = MsgBox("Main email recipient is missing. The email will not be sent.", vbCritical, "Error")
End If
End Sub
Function GetAllUserEmails() As String
Dim Rst As Variant
    GetAllUserEmails = ""
    Set Rst = New ADODB.Recordset
    Set Rst = CurrentDb.OpenRecordset("SELECT Tbl_Users.[E-mailAddress] FROM Tbl_Users WHERE (((Tbl_Users.[E-mailAddress]) Is Not Null));")
    If Rst.RecordCount > 0 Then
        Rst.MoveFirst
        While Not Rst.EOF
            GetAllUserEmails = GetAllUserEmails & "," & Rst.Fields("E-mailAddress")
            Rst.MoveNext
        Wend
        GetAllUserEmails = Mid(GetAllUserEmails, 2, Len(GetAllUserEmails))
    End If
    Set Rst = Nothing
End Function
