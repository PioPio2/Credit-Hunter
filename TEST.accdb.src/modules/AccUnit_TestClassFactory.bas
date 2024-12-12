Attribute VB_Name = "AccUnit_TestClassFactory"
Option Compare Text
Option Explicit
Option Private Module

Public Function AccUnitTestClassFactory_ZclsTestMultipleStatements() As Object
   Set AccUnitTestClassFactory_ZclsTestMultipleStatements = New ZclsTestMultipleStatements
End Function

Public Function AccUnitTestClassFactory_ZclsTestOutlook() As Object
   Set AccUnitTestClassFactory_ZclsTestOutlook = New ZclsTestOutlook
End Function

Public Function AccUnitTestClassFactory_ZclsTestUpdate() As Object
   Set AccUnitTestClassFactory_ZclsTestUpdate = New ZclsTestUpdate
End Function

Public Function AccUnitTestClassFactory_zclsTestUpdateFiles() As Object
   Set AccUnitTestClassFactory_zclsTestUpdateFiles = New zclsTestUpdateFiles
End Function
