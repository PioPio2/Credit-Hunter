Attribute VB_Name = "zTestDeleteTables"
Option Compare Database
Option Explicit

Sub DeleteTables()
'currentdb.Execute "sql....."
CurrentDb.Execute "DELETE Tbl_Customers.* FROM Tbl_Customers;"
CurrentDb.Execute "DELETE Tbl_Users.* FROM Tbl_Users;"

End Sub
