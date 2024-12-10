VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsUpdateFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Coll As Collection
Private SourceColl As Collection
Private DestinationColl As Collection
Private MainPathLen As Long
Private MainPath As String
Private NFilesCopied As Long
Public Function Scan(SourcePath, DestinationPath As String) As Integer
    Dim fso As Object    'FileSystemObject
    Set fso = CreateObject("scripting.FileSystemObject")    ' late binding

' scan destination folder
    Dim fldStart As Object    'Folder
    Set fldStart = fso.GetFolder(SourcePath)
    Dim Mask As String
    Mask = "*.*"
    Dim fld As Object    'Folder
    MainPath = SourcePath
    MainPathLen = Len(SourcePath)
    Set Coll = New Collection
    ListFiles fldStart, Mask
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next
    Set SourceColl = Coll

    ' scan destination folder
    Set Coll = New Collection
    Set fldStart = fso.GetFolder(DestinationPath)
    MainPath = DestinationPath
    MainPathLen = Len(DestinationPath)
    Mask = "*.*"
    ListFiles fldStart, Mask
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next
    Set DestinationColl = Coll
End Function
Public Function GetSourceList() As Collection
    Set GetSourceList = SourceColl
End Function
Public Function ListFolders(fldStart As Object, Mask As String)
    Dim fld As Object    'Folder
    Dim clsFile As zclsFile
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next
End Function

Public Sub ListFiles(fld As Object, Mask As String)
    Dim fl As Object    'File
    Dim clsFile As zclsFile
    For Each fl In fld.Files
        If fl.Name Like Mask Then
            Set clsFile = New zclsFile
            clsFile.Path = Left(MainPath, MainPathLen - 1)    'fld.Path & "\"
            clsFile.DateTime = fl.DateLastModified
            clsFile.FileName = Mid(fld.Path & "\" & fl.Name, MainPathLen, Len(fld.Path & "\" & fl.Name))
            Coll.Add clsFile
            Debug.Print fld.Path & "\" & fl.Name
        End If
    Next
End Sub
Public Function GetDestinationList() As Collection
    Set GetDestinationList = DestinationColl
End Function

Public Function Replace(ByRef ErrorN As Long, ByRef ErrorStr As String, ByRef NFiles As Long) As Boolean
    ErrorN = 0
    ErrorStr = ""
    NFiles = 0
    NFilesCopied = 0

    If SourceColl Is Nothing Then    ' if there is no sourcecoll created then error
        ErrorN = 1
        ErrorStr = "SourceColl was not created"
        Exit Function
    End If

    If SourceColl.Count = 0 Then    ' if there are no files in source then error
        ErrorN = 2
        ErrorStr = "There are is no file in the source folder"
        Exit Function
    End If

    ' scan source folder and compare with destination folder
    ' if any files in source is more recent than the correspondent file in destination then override the one in destination folder
    Dim SourceCount As Long
    Dim SourceFile As String
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    For SourceCount = 1 To SourceColl.Count    ' scan all files in source
        SourceFile = SourceColl.item(SourceCount).FileName
        Dim DestCount As Long
        Dim DestFile As String
        Dim OverrideFile As Boolean
        OverrideFile = False
        Dim FullSourcePath As String
        Dim FullDestinationPath As String
        For DestCount = 1 To DestinationColl.Count    ' scan destination files
            DestFile = DestinationColl.item(DestCount).FileName
            If DestFile = SourceFile Then    ' if file found in destination then check the datetime
                Dim SourceTime As Variant
                Dim DestTime As Variant
                SourceTime = SourceColl.item(SourceCount).DateTime
                DestTime = DestinationColl.item(DestCount).DateTime

                ' if file last modified date/time is more recent then the correspondent file in destination then override the file
                If SourceTime > DestTime Then
                    FullSourcePath = SourceColl.item(SourceCount).Path & SourceFile
                    FullDestinationPath = DestinationColl.item(DestCount).Path & DestFile
                    OverrideFile = True
                    Exit For
                Else
                    Exit For
                End If
            End If

            ' actioned only if there is a file in source that does not exist in destination folder
            If DestFile <> SourceFile And DestCount = DestinationColl.Count Then
                FullSourcePath = SourceColl.item(SourceCount).Path & SourceFile
                FullDestinationPath = DestinationColl.item(DestCount).Path & SourceFile
                OverrideFile = True
            End If
        Next DestCount

        ' the file in sourcefolder has to override the one in destination folder
        If OverrideFile = True Then
            Call fso.CopyFile(FullSourcePath, FullDestinationPath, True)
            NFilesCopied = NFilesCopied + 1
            NFiles = NFilesCopied
            OverrideFile = False
        End If

    Next SourceCount
End Function
