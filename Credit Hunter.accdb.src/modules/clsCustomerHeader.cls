VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Header As Scripting.Dictionary

Function Populate(aCustomerID As String) As Boolean
    Dim Support As ICustomerHeaderSupport
    Set Support = New clsCustomerHeaderSupportLIVE
    Set Header = Support.Populate(aCustomerID)
    Populate = True
End Function

Function getCustomerName() As String
    getCustomerName = Header("Name")
End Function

Function getCreditLimit() As String
    getCreditLimit = Header("Credit Limit")
End Function

Function getEmail() As String
    getEmail = Header("Email")
End Function

Function getAddress() As String
    If Not IsNull(Header("Address")) Then
        getAddress = Header("Address")
    Else
        getAddress = ""
    End If
End Function

Function getCustomerID() As String
    getCustomerID = Header("Customer_code")
End Function

Public Function GetHeader() As Scripting.Dictionary
    Set GetHeader = Header
End Function
