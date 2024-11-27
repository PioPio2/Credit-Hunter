﻿Attribute VB_Name = "AccUnit_Factory"
Option Compare Text
Option Explicit

#Const USE_ACCUNIT_TYPELIB = 1

Public Enum TestReportOutput
   DebugPrint = 1
   LogFile = 2
End Enum

#If USE_ACCUNIT_TYPELIB Then
#Else
Public Enum StringCompareMode
    StringCompareMode_BinaryCompare = 0
    StringCompareMode_TextCompare = 1
    StringCompareMode_vbNullStringEqualEmptyString = 4
End Enum
Public Enum ResetMode
    ResetMode_RemoveTests = 2
    ResetMode_ResetTestSuite = 4
End Enum
#End If

Private Const DefaultTestReportOutput As Long = TestReportOutput.DebugPrint
Private m_AccUnitLoaderFactory As Object
Private m_CodeCoverageTracker As Object

Private Function AccUnitLoaderFactory() As Object
   If m_AccUnitLoaderFactory Is Nothing Then
      Set m_AccUnitLoaderFactory = GetAccUnitLoaderFactory
   End If
   Set AccUnitLoaderFactory = m_AccUnitLoaderFactory
End Function

Private Function GetAccUnitLoaderFactory() As Object

   Dim AccUnitVbeAddIn As Object

   If TryGetAccUnitVbeAddIn(AccUnitVbeAddIn) Then
      Set GetAccUnitLoaderFactory = AccUnitVbeAddIn.Object
   Else
      Set GetAccUnitLoaderFactory = Application.Run(GetAddInPath & "AccUnitLoader.GetAccUnitFactory")
   End If

End Function

Private Function TryGetAccUnitVbeAddIn(ByRef AccUnitVbeAddIn As Object) As Boolean

   Dim AddIn2check As Object

   For Each AddIn2check In Application.VBE.AddIns
      If AddIn2check.progID = "AccUnit.VbeAddIn.Connect" Then
         If AddIn2check.Connect Then
            Set AccUnitVbeAddIn = Application.VBE.AddIns.item("AccUnit.VbeAddIn.Connect")
            TryGetAccUnitVbeAddIn = True
         End If
      End If
   Next

End Function

#If USE_ACCUNIT_TYPELIB Then
Private Property Get AccUnitFactory() As AccUnit.AccUnitFactory
#Else
Private Property Get AccUnitFactory() As Object
#End If
   Set AccUnitFactory = AccUnitLoaderFactory.AccUnitFactory
End Property

Private Function GetAddInPath() As String
   GetAddInPath = Environ("appdata") & "\Microsoft\AddIns\"
End Function

#If USE_ACCUNIT_TYPELIB Then
Public Property Get Assert() As AccUnit.Assert
#Else
Public Property Get Assert() As Object
#End If
   Set Assert = AccUnitLoaderFactory.Assert
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Property Get Iz() As AccUnit.ConstraintBuilder
#Else
Public Property Get Iz() As Object
#End If
    Set Iz = AccUnitLoaderFactory.ConstraintBuilder
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Property Get TestSuite(Optional ByVal OutputTo As TestReportOutput = DefaultTestReportOutput) As AccUnit.AccessTestSuite
#Else
Public Property Get TestSuite(Optional ByVal OutputTo As TestReportOutput = DefaultTestReportOutput) As Object
#End If
   Set TestSuite = AccUnitLoaderFactory.TestSuite(OutputTo)
   TestSuite.Reset ResetMode_ResetTestSuite + ResetMode_RemoveTests
End Property

Public Sub RunAllTests()
   TestSuite.AddFromVBProject.Run
End Sub

#If USE_ACCUNIT_TYPELIB Then
Public Property Get CodeCoverageTracker(Optional ReInit As Boolean = False) As AccUnit.CodeCoverageTracker
#Else
Public Property Get CodeCoverageTracker(Optional ReInit As Boolean = False) As Object
#End If
   If ReInit Then
      If Not m_CodeCoverageTracker Is Nothing Then
         m_CodeCoverageTracker.Dispose
         Set m_CodeCoverageTracker = Nothing
      End If
   End If
   If m_CodeCoverageTracker Is Nothing Then
      Set m_CodeCoverageTracker = AccUnitLoaderFactory.CodeCoverageTracker
   End If
   Set CodeCoverageTracker = m_CodeCoverageTracker
End Property

#If USE_ACCUNIT_TYPELIB Then
Public Function CodeCoverageTest(ParamArray CodeModulNames() As Variant) As AccUnit.AccessTestSuite
#Else
Public Function CodeCoverageTest(ParamArray CodeModulNames() As Variant) As Object
#End If
   Dim CodeModuleName As Variant
   Dim CodeCoverageTestSuite As Object

   With CodeCoverageTracker(True)
      For Each CodeModuleName In CodeModulNames
         .Add CodeModuleName
      Next
   End With

   Set CodeCoverageTestSuite = AccUnitLoaderFactory.TestSuite(DefaultTestReportOutput)
   Set CodeCoverageTestSuite.CodeCoverageTracker = m_CodeCoverageTracker

   Set CodeCoverageTest = CodeCoverageTestSuite

End Function

#If USE_ACCUNIT_TYPELIB Then
Public Property Get ErrorTrappingObserver() As AccUnit.AccessErrorTrappingObserver
#Else
Public Property Get ErrorTrappingObserver() As Object
#End If
   Set ErrorTrappingObserver = AccUnitLoaderFactory.AccessErrorTrappingObserver()
End Property
