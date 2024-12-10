VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zclsTestUpdateFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const SourcePath = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Source\"
Const DestinationPath = "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Destination\"
'AccUnit:TestClass

Public Sub Setup()
' This procedure is called before each test
    Call DeleteTables
    On Error Resume Next
    Kill (SourcePath & "impostor.txt")
    Kill (DestinationPath & "impostor.txt")
    On Error GoTo 0
End Sub

Public Sub Teardown()
' This procedure is called after each test
End Sub

Public Sub TestScan()

    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles
    Call CollectFiles.Scan("C:\Users\alberto\OneDrive\Credit Hunter\Tests\Source\", "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Destination\")

    Dim Actual As Variant
    Actual = CollectFiles.GetSourceList.Count
    Assert.That Actual, Iz.EqualTo(3)

    Actual = CollectFiles.GetDestinationList.Count
    Assert.That Actual, Iz.EqualTo(3)

    Dim SL As Collection
    Set SL = CollectFiles.GetSourceList
    Dim DL As Collection
    Set DL = CollectFiles.GetDestinationList
End Sub

Public Sub TestReplaceNoColl()
    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles

    Dim Actual As Variant
    Dim Errors As String
    Dim ErrorN As Long
    Dim NFiles As Long
    Assert.IsFalse (CollectFiles.Replace(ErrorN, Errors, NFiles))
    Assert.AreEqual 1, ErrorN
End Sub

Public Sub TestReplaceColl0()
    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles
    Call CollectFiles.Scan("C:\Users\alberto\OneDrive\Credit Hunter\Tests\EmptyFolder", "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Destination\")

    Dim Actual As Variant
    Dim Errors As String
    Dim ErrorN As Long
    Dim NFiles As Long
    Assert.IsFalse (CollectFiles.Replace(ErrorN, Errors, NFiles))
    Assert.AreEqual 2, ErrorN
End Sub
Public Sub TestReplace1Fileupdated()
    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles

    ' Set last modified datetime of Test.docs to now() in order to make this file only to be updated

    Const FileTochange = "Test.docx"
    Dim ST As Double
    ST = Now()
    Dim Res As Boolean
    Res = SetFileDateTime(SourcePath & FileTochange, ST, FileDateLastModified, False)


    Call CollectFiles.Scan(SourcePath, "C:\Users\alberto\OneDrive\Credit Hunter\Tests\Destination\")

    Dim Actual As Variant
    Dim Errors As String
    Dim ErrorN As Long
    Dim NFiles As Long
    Assert.IsFalse (CollectFiles.Replace(ErrorN, Errors, NFiles))
    Assert.AreEqual 0, ErrorN    'no error
    Assert.AreEqual 1, NFiles    ' one file updated - Test.docx
End Sub





Public Sub TestReplace2Fileupdated()
    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles

    ' Set last modified datetime of Test.docs to now() in order to make this file only to be updated

    Const FileTochange = "Test.docx"
    Dim ST As Double
    ST = Now()
    Dim Res As Boolean
    Res = SetFileDateTime(SourcePath & FileTochange, ST, FileDateLastModified, False)

    ' create additional file only for this test only
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim a As Variant
    Set a = fs.CreateTextFile(SourcePath & "impostor.txt", True)
    a.WriteLine ("This is a test.")
    a.Close

    Call CollectFiles.Scan(SourcePath, DestinationPath)

    Dim Actual As Variant
    Dim Errors As String
    Dim ErrorN As Long
    Dim NFiles As Long
    Assert.IsFalse (CollectFiles.Replace(ErrorN, Errors, NFiles))
    Assert.AreEqual 0, ErrorN    'no error
    Assert.AreEqual 1, NFiles    ' one file updated - Test.docx
End Sub
