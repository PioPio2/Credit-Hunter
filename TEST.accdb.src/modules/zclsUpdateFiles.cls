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
' how to change datetime of a file
'    Dim WhatTime As FileDateToProcess
'    Dim result As Boolean
'    result = SetFileDateTime(DestinationColl.item(1).Path & DestinationColl.item(1).FileName, SourceColl.item(1).DateTime, _
'                             WhichDateToChange:=FileDateLastModified, NoGMTConvert:=False)
End Function
Public Function GetSourceList() As Collection
    Set GetSourceList = SourceColl
End Function
Public Function ListFolders(fldStart As Object, Mask As String)
    Dim fld As Object    'Folder
    Dim clsFile As zclsFile
    For Each fld In fldStart.SubFolders
        '        Set clsFile = New zclsFile
        '        clsFile.Filename = fld.Path & "\"
        '        clsFile.DateTime = fld.DateLastModified
        '        Coll.Add clsFile
        '        Debug.Print fld.Path & "\"
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
            clsFile.DateTime = fld.DateLastModified
            clsFile.FileName = Mid(fld.Path & "\" & fl.Name, MainPathLen, Len(fld.Path & "\" & fl.Name))
            Coll.Add clsFile
            Debug.Print fld.Path & "\" & fl.Name
        End If
    Next
End Sub
Public Function GetDestinationList() As Collection
    Set GetDestinationList = DestinationColl
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Call the SetFileTimeLastModified flavor of SetFileTime.
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim Res As Long
'Res = SetFileTimeLastModified( _
  '          hFile:=FileHandle, _
  '          CreateTime:=NULL_LONG, _
  '          LastAccessTime:=NULL_LONG, _
  '          LastModified:=tFileTime)
'If Res = 0 Then
'ErrNum = Err.LastDllError
'ErrText = GetSystemErrorMessageText(ErrNum)
'Debug.Print "Error With SetFileTimeLastModified:" & vbCrLf & _
'            "Err:  " & CStr(ErrNum) & vbCrLf & _
'            "Desc: " & ErrText
'
