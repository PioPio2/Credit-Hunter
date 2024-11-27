VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iStatementSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function GetTemplatePath() As String

End Function

Public Function OpenStatementTemplate(PathTemplate As String) As Boolean

End Function
Public Function StatementVisible(Visible As Boolean) As Boolean

End Function
Public Function StatementReplace(aWhat As String, aWith As String) As Boolean

End Function

Public Function CopyStatementData(aCellStart As Range, aData As ADODB.Recordset, ByRef outOutstanding As Currency, ByRef outOverdue As Currency, ByVal NRec As Long) As Long

End Function
Public Function FindStatementDataStart() As Range

End Function
Public Function SaveStatement(CompleteFileName As String) As Boolean

End Function
Public Function CloseStatement() As Boolean

End Function
Public Function GetSQLStatement(CustomerCode As String) As String

End Function
Public Function GetRecordsetStatement(SQLString As String) As ADODB.Recordset

End Function
Public Function GetDataStatement(CustomerID As String, ByRef NRec As Long) As ADODB.Recordset

End Function
Public Function GetStatementSavePath() As String

End Function
Public Sub Maximize()

End Sub
