VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelinkDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function ReLinkDB() As Boolean
Dim db As DAO.Database
    Dim app As Access.Application
    Dim wks As Workspace
    Dim tdf As DAO.TableDef
    Dim newConnectString As String
    Dim tableName As String

    On Error GoTo ErrorHandler

    Dim Support  As clsRelinkDBSupportLIVE
    Set Support = New clsRelinkDBSupportLIVE
    Dim NewBackEndPath As String
    NewBackEndPath = Support.GetDBPath ' retrieve folder where the production db is in order to re-link the tables

    ' Get the current database
    Set db = CurrentDb()
'    Set app = New Access.Application
 '   Set wks = app.DBEngine(0)
  '  Set db = wks.OpenDatabase("E:\MS Access\Projects\Credit Hunter\Credit Hunter.accdb")
    ' Loop through each TableDef in the database
    For Each tdf In db.TableDefs
        ' Check if the table is a linked table (not a local table)
        If Len(tdf.Connect) > 0 Then
            ' Update the Connect property with the new backend path
            tableName = tdf.Name
            newConnectString = "MS Access;DATABASE=" & NewBackEndPath
            tdf.Connect = newConnectString

            ' Refresh the link to the table
            tdf.RefreshLink

           ' MsgBox "Linked table '" & tableName & "' successfully updated."
        End If
    Next tdf
    MsgBox ("Relink terminated succesfully.")

ExitSub:
    ' Cleanup
    Set tdf = Nothing
    Set db = Nothing
    ReLinkDB = True
    Exit Function

ErrorHandler:
    MsgBox "Error updating linked tables: " & err.Description
    Resume ExitSub

End Function
