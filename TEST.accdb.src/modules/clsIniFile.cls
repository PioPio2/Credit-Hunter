VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function IniFileName() As String
Dim CurrentPath As String
    CurrentPath = Application.CurrentProject.Path
    IniFileName = CurrentPath & "\Credit Hunter.ini"
End Function


Public Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
Dim Worked As Long
Dim RetStr As String * 128
Dim StrSize As Long
Dim iNoOfCharInIni As Long
Dim sIniString As String
Dim sProfileString As String

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

Public Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
Dim Worked As Long
Dim iNoOfCharInIni As Long
Dim sIniString As String
Dim sProfileString As String

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
End Function
