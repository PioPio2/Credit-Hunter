Attribute VB_Name = "zTestDeleteTables"
Option Compare Database
Option Explicit

Sub DeleteTables()
'currentdb.Execute "sql....."
    CurrentDb.Execute "DELETE Tbl_Customers.* FROM Tbl_Customers;"
    CurrentDb.Execute "DELETE Tbl_Users.* FROM Tbl_Users;"
    CurrentDb.Execute "DELETE Tbl_CustomersList.* FROM Tbl_CustomersList;"
    CurrentDb.Execute "DELETE Tbl_Invoices.* FROM Tbl_Invoices;"
    CurrentDb.Execute "DELETE Tbl_Customers.* FROM Tbl_Customers;"
End Sub
