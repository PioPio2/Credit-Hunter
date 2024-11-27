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

Private Function IsFile(ByVal fName As String) As Boolean
    'Returns TRUE if the provided name points to an existing file.
    'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Public Function DownloadFileFromGitHub(GitRawFile, UpdateDestinationFile As String) As Boolean
    Dim http As Object
    Dim filePath As String
    Dim url As String
    Dim fileStream As Object
    Dim byteArray() As Byte
    
    ' URL of the raw file from GitHub
    'url = "https://bitbucket.org/niftyanswer/credit-hunter/raw/b035250a7448299e3a5b4b791b7ede51b8d201db/LatestUpdate/LatestUpdate.txt"
    'url = "https://bitbucket.org/niftyanswer/credit-hunter/raw/b035250a7448299e3a5b4b791b7ede51b8d201db/Credit%20Hunter.accdb" appliaction raw file  location
    url = GitRawFile
    
    ' Define the local file path where the file will be saved
    'filePath = "E:\MS Access\Projects\Credit Hunter\LatestUpdate.txt"
    'filePath = "E:\MS Access\Projects\Credit Hunter\TEST update\TEST.accdb"
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

Public Function LatestVersion(GitRawFile, FileName As String) As Variant
    If IsFile(FileName) = False Then
        LatestVersion = 0
    Else
        Dim LatestVersionFond As Variant
        LatestVersionFond = DownloadFileFromGitHub(GitRawFile, FileName)
        If LatestVersionFond = True Then
            Dim strIn As String
            Dim objFSO As Object
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Dim objTF As Object
            Set objTF = objFSO.OpenTextFile(FileName, 1)
            strIn = objTF.readall
            strIn = Trim(strIn)
            LatestVersion = CDec(strIn)
            objTF.Close
        Else
            LatestVersion = 0
        End If
    End If
End Function

Public Function UpdateLocalReleaseN(RelN As Variant) As Boolean
    Dim SQL As String
    SQL = "UPDATE Release SET Release.rel = " & RelN & ", Release.CodeRelease = " & RelN & ", Release.ReleaseDate = #" & Format(Date, "mm/dd/yyyy") & "#;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL SQL
    DoCmd.SetWarnings True
    UpdateLocalReleaseN = True
End Function
