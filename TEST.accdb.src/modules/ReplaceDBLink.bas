Attribute VB_Name = "ReplaceDBLink"
Option Compare Database
Option Explicit

Sub RelinkTables(NewBackEndPath As String)
Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim newConnectString As String
    Dim tableName As String

    On Error GoTo ErrorHandler

    ' Get the current database
    Set db = CurrentDb()

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

'    MsgBox "All linked tables have been successfully updated to the new backend."

ExitSub:
    ' Cleanup
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error updating linked tables: " & Err.Description
    Resume ExitSub
End Sub
