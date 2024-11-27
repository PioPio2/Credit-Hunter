VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Support As clsUsersSupport

Public Function CurrentUserEmailAddress() As String
    CreateSupport
    CurrentUserEmailAddress = Support.CurrentUserEmailAddress
    Set Support = Nothing
End Function
Private Sub CreateSupport()
    If Support Is Nothing Then
        Set Support = New clsUsersSupport
    End If
End Sub
