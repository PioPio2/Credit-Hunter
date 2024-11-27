VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUsersSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private RS As DAO.Recordset
Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetCurrentUserName() As String
' Returns the network login name
    Dim lngLen As Long, lngX As Long
    Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If lngX <> 0 Then
        GetCurrentUserName = Left$(strUserName, lngLen - 1)
    Else
        GetCurrentUserName = ""
    End If
End Function
Private Sub CreateRS()
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "        [Tbl_Users].* "
    SQL = SQL & "From "
    SQL = SQL & "        [Tbl_Users] "
    SQL = SQL & "Where "
    SQL = SQL & "        [Tbl_Users].UserName = '" & GetCurrentUserName & "' "
    Set RS = CurrentDb.OpenRecordset(SQL)
End Sub

Public Function CurrentUserEmailAddress() As String
    Call CreateRS
    CurrentUserEmailAddress = RS.Fields("E-mailAddress").value
End Function
