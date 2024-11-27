Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Load()
'  DoCmd.OpenForm "mskwhopaidyesterday2"
' Application.Forms.item("mskwhopaidyesterday2").SetFocus
End Sub

Private Sub Form_Open(Cancel As Integer)
' Create Outlook instance that can be used everywhere in the application. This is to increase performance
' Set GeneralOL = New Outlook.Application
' Set GeneralWD = New Word.Application

    Set OL = New clsOutlook
    OL.CreateOutlook

    ' Minimize the database window and initialize the form.

    ' Move to the switchboard page that is marked as the default.
    Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
    Me.FilterOn = True


    Dim LocalReleaseNumber As Variant
    Dim IniFile As clsIniFile
    Set IniFile = New clsIniFile
    LocalReleaseNumber = CDec(IniFile.ReadIniFileString("Local Version", "Ver"))

    Dim Update As clsUpdate
    Set Update = New clsUpdate

    Dim GitLatestRelease As Variant
    GitLatestRelease = DLookup("GitLatestReleaseNumberFile", "TblGeneral")

    Dim LocalLatestReleasePath As String
    LocalLatestReleasePath = DLookup("LocalLatestReleaseNumberFile", "TblGeneral")

    Dim LocalLatestReleaseFile As String
    LocalLatestReleaseFile = LocalLatestReleasePath & "LatestUpdate.txt"
    Dim LatestReleaseNumber As Variant
    LatestReleaseNumber = Update.LatestVersion(GitLatestRelease, LocalLatestReleaseFile)    ' get latest release n# available online

    If LatestReleaseNumber > LocalReleaseNumber Then    ' if there is an update then download it
        MsgBox ("There is a new version of Credit Hunter. I will download now.")
        Dim UpdateDownloaded As String
        UpdateDownloaded = LocalLatestReleasePath & "Credit Hunter.accdb"
        Dim GitLatestUpdatedApp As String
        GitLatestUpdatedApp = DLookup("GitUpdatedApp", "TblGeneral")
        Call Update.DownloadFileFromGitHub(GitLatestUpdatedApp, UpdateDownloaded)

        ' launch update
        Dim wsShell As Object
        Set wsShell = CreateObject("WScript.Shell")
        'wsShell.Run Chr(34) & Application.CurrentProject.path & "\Update.accdb" & Chr(34) reactivate later
        MsgBox ("Close the application and activate the update manually")
        Set wsShell = Nothing
    End If

    Dim Support As IWhoPaidYesterday
    Set Support = New clsWhoPaidYesterday
    Dim Injection As IWhoPaidYesterdaySupport
    Set Injection = New clsWhoPaidYesterdaySupportLIVE
    Call Support.Inject(Injection)

    Call Support.UpdateSecondaryQueries          ' update two queries that retrieve the customer who paid something from the last import

    Dim SQLString As String
    SQLString = Support.GetRecordSourceSQL
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQLString)
    If RS.RecordCount = 0 Then                   ' if nobody paid from the last report then exit Sub
        Exit Sub
    End If
End Sub

Private Sub Form_Current()
' Update the caption and fill in the list of options.

    Me.Caption = Nz(Me![ItemText], "")
    FillOptions

End Sub

Private Sub FillOptions()
' Fill in the options for this switchboard page.

' The number of buttons on the form.
    Const conNumButtons = 8

    Dim con As Object
    Dim RS As Object
    Dim stSql As String
    Dim intOption As Integer

    ' Set the focus to the first button on the form,
    ' and then hide all of the buttons on the form
    ' but the first.  You can't hide the field with the focus.
    Me![Option1].SetFocus
    For intOption = 2 To conNumButtons
        Me("Option" & intOption).Visible = False
        Me("OptionLabel" & intOption).Visible = False
    Next intOption

    ' Open the table of Switchboard Items, and find
    ' the first item for this Switchboard Page.
    Set con = Application.CurrentProject.Connection
    stSql = "SELECT * FROM [Switchboard Items]"
    stSql = stSql & " WHERE [ItemNumber] > 0 AND [SwitchboardID]=" & Me![SwitchboardID]
    stSql = stSql & " ORDER BY [ItemNumber];"
    Set RS = CreateObject("ADODB.Recordset")
    RS.Open stSql, con, 1                        ' 1 = adOpenKeyset

    ' If there are no options for this Switchboard Page,
    ' display a message.  Otherwise, fill the page with the items.
    If (RS.EOF) Then
        Me![OptionLabel1].Caption = "There are no items for this switchboard page"
    Else
        While (Not (RS.EOF))
            Me("Option" & RS![ItemNumber]).Visible = True
            Me("OptionLabel" & RS![ItemNumber]).Visible = True
            Me("OptionLabel" & RS![ItemNumber]).Caption = RS![ItemText]
            RS.MoveNext
        Wend
    End If

    ' Close the recordset and the database.
    RS.Close
    Set RS = Nothing
    Set con = Nothing

End Sub

Private Function HandleButtonClick(intBtn As Integer)
' This function is called when a button is clicked.
' intBtn indicates which button was clicked.

' Constants for the commands that can be executed.
    Const conCmdGotoSwitchboard = 1
    Const conCmdOpenFormAdd = 2
    Const conCmdOpenFormBrowse = 3
    Const conCmdOpenReport = 4
    Const conCmdCustomizeSwitchboard = 5
    Const conCmdExitApplication = 6
    Const conCmdRunMacro = 7
    Const conCmdRunCode = 8
    Const conCmdOpenPage = 9

    ' An error that is special cased.
    Const conErrDoCmdCancelled = 2501

    Dim con As Object
    Dim RS As Object
    Dim stSql As String

    On Error GoTo HandleButtonClick_Err

    ' Find the item in the Switchboard Items table
    ' that corresponds to the button that was clicked.
    Set con = Application.CurrentProject.Connection
    Set RS = CreateObject("ADODB.Recordset")
    stSql = "SELECT * FROM [Switchboard Items] "
    stSql = stSql & "WHERE [SwitchboardID]=" & Me![SwitchboardID] & " AND [ItemNumber]=" & intBtn
    RS.Open stSql, con, 1                        ' 1 = adOpenKeyset

    ' If no item matches, report the error and exit the function.
    If (RS.EOF) Then
        MsgBox "There was an error reading the Switchboard Items table."
        RS.Close
        Set RS = Nothing
        Set con = Nothing
        Exit Function
    End If

    Select Case RS![Command]

        ' Go to another switchboard.
    Case conCmdGotoSwitchboard
        Me.Filter = "[ItemNumber] = 0 AND [SwitchboardID]=" & RS![Argument]

        ' Open a form in Add mode.
    Case conCmdOpenFormAdd
        DoCmd.OpenForm RS![Argument], , , , acAdd

        ' Open a form.
    Case conCmdOpenFormBrowse
        DoCmd.OpenForm RS![Argument]

        ' Open a report.
    Case conCmdOpenReport
        DoCmd.OpenReport RS![Argument], acPreview

        ' Customize the Switchboard.
    Case conCmdCustomizeSwitchboard
        ' Handle the case where the Switchboard Manager
        ' is not installed (e.g. Minimal Install).
        On Error Resume Next
        Application.Run "ACWZMAIN.sbm_Entry"
        If (err <> 0) Then MsgBox "Command not available."
        On Error GoTo 0
        ' Update the form.
        Me.Filter = "[ItemNumber] = 0 AND [Argument] = 'Default' "
        Me.Caption = Nz(Me![ItemText], "")
        FillOptions

        ' Exit the application.
    Case conCmdExitApplication
        CloseCurrentDatabase

        ' Run a macro.
    Case conCmdRunMacro
        DoCmd.RunMacro RS![Argument]

        ' Run code.
    Case conCmdRunCode
        Application.Run RS![Argument]

        ' Open a Data Access Page
    Case conCmdOpenPage
        DoCmd.OpenDataAccessPage RS![Argument]

        ' Any other command is unrecognized.
    Case Else
        MsgBox "Unknown option."

    End Select

    ' Close the recordset and the database.
    RS.Close

HandleButtonClick_Exit:
    On Error Resume Next
    Set RS = Nothing
    Set con = Nothing
    Exit Function

HandleButtonClick_Err:
    ' If the action was cancelled by the user for
    ' some reason, don't display an error message.
    ' Instead, resume on the next line.
    If (err = conErrDoCmdCancelled) Then
        Resume Next
    Else
        MsgBox "There was an error executing the command.", vbCritical
        Resume HandleButtonClick_Exit
    End If

End Function

Private Sub Form_Timer()
    Me.TimerInterval = 0
    DoCmd.OpenForm "mskwhopaidyesterday2"
    On Error Resume Next
    Application.Forms.item("mskwhopaidyesterday2").SetFocus
    On Error GoTo 0
End Sub
