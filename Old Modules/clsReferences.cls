VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub CheckReferenceRubberDuck()
    Dim ref As Reference
    Dim refStatus As String

    ' Loop through each reference in the References collection
    For Each ref In Application.References
        ' Check if the reference is broken (missing)
        If ref.IsBroken Then
            refStatus = "MISSING"
        Else
            refStatus = "Available"
        End If

        ' Print the name and status of each reference
        Debug.Print "Reference: " & ref.Name & " - " & refStatus
    Next ref
End Sub

Sub AddReferenceRubberDuck()
' Adds a reference to the Microsoft Scripting Runtime library
    On Error Resume Next
    References.AddFromFile "C:\ProgramData\Rubberduck\Rubberduck.x32.tlb"
    On Error GoTo 0
End Sub
Sub RemoveReferenceRubberduck()
    Dim ref As Reference
    ' Loop through the references to find the one to remove
    For Each ref In References
        If ref.Name = "Rubberduck" Then    ' Replace "Scripting" with the name of the reference
            References.Remove ref
            Exit For
        End If
    Next ref
End Sub
