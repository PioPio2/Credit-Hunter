VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZclsTestUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Const SourcePath = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Source\"
'Const DestinationPath = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Destination\"
'AccUnit:TestClass

Public Sub Setup()
' This procedure is called before each test
'    Call DeleteTables
'    On Error Resume Next
'    Kill (SourcePath & "impostor.txt")
'    Kill (DestinationPath & "impostor.txt")
'    On Error GoTo 0
End Sub

Public Sub Teardown()
' This procedure is called after each test
End Sub

Public Sub TestScanFalse()
' set rel=10
    Dim Inifile As clsIniFile
    Set Inifile = New clsIniFile
    Inifile.SetFileName ("Test.ini")
    Call Inifile.WriteIniFileString("Local Version", "Ver", "10")

    Dim Update As clsUpdate
    Set Update = New clsUpdate
    Dim Actual As Variant
    Actual = Update.Scan("Test.ini")
    Assert.IsFalse Actual
End Sub

Public Sub TestScanTrue()
' set rel=1
    Dim Inifile As clsIniFile
    Set Inifile = New clsIniFile
    Inifile.SetFileName ("Test.ini")
    Call Inifile.WriteIniFileString("Local Version", "Ver", "1")

    Dim Update As clsUpdate
    Set Update = New clsUpdate
    Dim Actual As Variant
    Actual = Update.Scan("Test.ini")
    Assert.IsTrue Actual
End Sub

Public Sub TestDownloadTrue()

    Dim Update As clsUpdate
    Set Update = New clsUpdate

    Const url = "https://github.com/PioPio2/Credit-Hunter/blob/main/README.md"
    Const Destination = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\README.md"
    On Error Resume Next
    Kill (Destination)
    On Error GoTo 0

    Call Update.DownloadFileFromGitHub(url, Destination)
    Dim Actual As Boolean
    Actual = ((GetAttr(Destination) And vbDirectory) <> vbDirectory)
    Assert.IsTrue Actual
End Sub
Public Sub TestDownloadfalse()

    Dim Update As clsUpdate
    Set Update = New clsUpdate

    Const url = "https://github.com/PioPio2/Credit-Hunter/blob/main/README.md"
    Const Destination = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\README.md"
    On Error Resume Next
    Kill (Destination)
    On Error GoTo 0

    Call Update.DownloadFileFromGitHub(url, Destination)
    On Error Resume Next
    Kill Destination
    Dim Actual As Boolean
    Actual = ((GetAttr(Destination) And vbDirectory) <> vbDirectory)
    Assert.False Actual
    On Error GoTo 0
End Sub
