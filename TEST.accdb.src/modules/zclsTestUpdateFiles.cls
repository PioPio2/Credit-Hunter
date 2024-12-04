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

'AccUnit:TestClass

Public Sub Setup()
' This procedure is called before each test
    Call DeleteTables
End Sub

Public Sub Teardown()
' This procedure is called after each test
End Sub

Public Sub TestScan()

    Dim CollectFiles As zclsUpdateFiles
    Set CollectFiles = New zclsUpdateFiles
    CollectFiles.Scan ("C:\Users\alberto\OneDrive\Credit Hunter\TEST.accdb.src\")

    Dim Actual As Variant
    Actual = CollectFiles.GetFileList.Count
    Assert.That Actual, Iz.GreaterThan(0)

End Sub
