VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerHeaderSupportLIVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Implements ICustomerHeaderSupport
Private Dict As Scripting.Dictionary

Public Function ICustomerHeaderSupport_GetDictionary() As Scripting.Dictionary
    Set ICustomerHeaderSupport_GetDictionary = Dict
End Function

Public Function ICustomerHeaderSupport_Populate(aCustomerID As String) As Scripting.Dictionary
    Dim SQL As String
    SQL = ""
    SQL = SQL & "Select "
    SQL = SQL & "    [Tbl_Customers].* "
    SQL = SQL & "From "
    SQL = SQL & "    [Tbl_Customers] "
    SQL = SQL & "Where "
    SQL = SQL & "    [Tbl_Customers].[Customer_code] = '" & aCustomerID & "' "
    Dim RS As DAO.Recordset
    Set RS = Application.CurrentDb.OpenRecordset(SQL)

    Set Dict = New Scripting.Dictionary
    Dim Fld As Variant
    If RS.RecordCount > 0 Then
        For Each Fld In RS.Fields
            Dict.Add Fld.Name, Fld.value
        Next
    End If
    Set ICustomerHeaderSupport_Populate = Dict
End Function
