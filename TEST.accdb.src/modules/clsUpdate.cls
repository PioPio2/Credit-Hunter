VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Function Scan(IniFileName As String) As Boolean
' find local rel n#
    Dim Ini As clsIniFile
    Set Ini = New clsIniFile
    Ini.SetFileName (IniFileName)
    Dim LocalVer As Double
    LocalVer = Ini.ReadIniFileString("Local Version", "Ver")

    ' find remote rel n#
    Dim OnlineRel As Double
    OnlineRel = LatestVersion
    Set Ini = Nothing

    'compare the two and if remote is most updated then flag it
    Scan = OnlineRel > LocalVer
End Function
Public Function Update() As Boolean

End Function
Public Function Restart() As Boolean

End Function
Private Function isfile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    isfile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Public Function DownloadFileFromGitHub(GitRawFile, UpdateDestinationFile As String) As Boolean
    Dim http As Object
    Dim filePath As String
    Dim url As String
    Dim fileStream As Object
    Dim byteArray() As Byte

    ' URL of the raw file from GitHub
    url = GitRawFile

    ' Define the local file path where the file will be saved
    filePath = UpdateDestinationFile

    ' Create XMLHTTP object for HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Send a GET request to GitHub to retrieve the file
    http.Open "GET", url, False
    http.Send

    ' Check if the request was successful (HTTP status code 200)
    If http.Status = 200 Then
        ' Get the response as a byte array (binary)
        byteArray = http.responseBody

        ' Create ADODB Stream to save the byte array to a file
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1                      ' Binary data
        fileStream.Open
        fileStream.Write byteArray

        ' Save the file to the specified file path
        fileStream.SaveToFile filePath, 2        ' 2 = overwrite if file exists

        ' Close the stream
        fileStream.Close
        DownloadFileFromGitHub = True
        'MsgBox "File downloaded successfully to " & filePath
    Else
        DownloadFileFromGitHub = False
        'MsgBox "Error: Unable to download the file. HTTP Status: " & http.Status
    End If
End Function

Public Function LatestVersion() As Variant
    Dim Whttp As Object
    Dim url As String
    Set Whttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "https://github.com/PioPio2/Credit-Hunter/raw/refs/heads/main/LatestUpdate/LatestUpdate.txt"   'SharePoint Path For The File

    Dim response As String
    With Whttp
        .Open "Get", url
        .Send
        LatestVersion = .responseText
    End With
End Function

Public Function UpdateLocalReleaseN(RelN As Variant) As Boolean
    Dim SQL As String
    SQL = "UPDATE Release SET Release.rel = " & RelN & ", Release.CodeRelease = " & RelN & ", Release.ReleaseDate = #" & Format(Date, "mm/dd/yyyy") & "#;"

    DoCmd.SetWarnings False
    DoCmd.RunSQL SQL
    DoCmd.SetWarnings True
    UpdateLocalReleaseN = True
End Function
