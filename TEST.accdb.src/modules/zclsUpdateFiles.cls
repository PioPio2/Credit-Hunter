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
Dim Coll As Collection
Public Function Scan(Path As String) As Integer
    'Dim Coll As Collection
    Set Coll = New Collection

    Dim fso As Object    'FileSystemObject
    Dim fldStart As Object    'Folder
    Dim fld As Object    'Folder
    Dim fl As Object    'File
    Dim Mask As String

    Set fso = CreateObject("scripting.FileSystemObject")    ' late binding
    'Set fso = New FileSystemObject 'or use early binding (also replace Object types)

    Set fldStart = fso.GetFolder(Path)

    Mask = "*.*"
    'Debug.Print fldStart.Path & "\"
    ListFiles fldStart, Mask
    For Each fld In fldStart.SubFolders
        ListFiles fld, Mask
        ListFolders fld, Mask
    Next

End Function
Public Function GetFileList() As Collection
    Set GetFileList = Coll
End Function
Public Function ListFolders(fldStart As Object, Mask As String)
    Dim fld As Object    'Folder
    Dim clsFile As zclsFile
    For Each fld In fldStart.SubFolders
        Set clsFile = New zclsFile
        clsFile.Filename = fld.Path & "\"
        clsFile.DateTime = fld.DateLastModified
        Coll.Add clsFile
        Debug.Print fld.Path & "\"
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
            clsFile.Filename = fld.Path & "\"
            clsFile.DateTime = fld.DateLastModified
            Coll.Add clsFile
            Debug.Print fld.Path & "\" & fl.Name
        End If
    Next
    End Sub

    '    Next
    '
    'End Sub
    '
End Sub
