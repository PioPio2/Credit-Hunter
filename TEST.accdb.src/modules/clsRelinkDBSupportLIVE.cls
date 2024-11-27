VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelinkDBSupportLIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function GetDBPath() As String
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        TblGeneral.PathDB "
    SQL = SQL & "From "
    SQL = SQL & "        TblGeneral "
    Dim RS As DAO.Recordset
    Set RS = CurrentDb.OpenRecordset(SQL)
    GetDBPath = RS.Fields(0).value
End Function
