Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Dim Inputdata As String
Dim StrInputFileName, StrFilter As String
StrFilter = ahtAddFilterItem(StrFilter, "Text file(*.txt)", "*.txt")
StrInputFileName = ahtCommonFileOpenSave(Filter:=StrFilter, OpenFile:=True, _
                DialogTitle:="Please select the file...", _
                Flags:=ahtOFN_HIDEREADONLY)
If StrInputFileName <> "" Then
    Open StrInputFileName For Input As #1 'check the header file to understand if the source file is correct
    Line Input #1, Inputdata
    If Inputdata <> DLookup("CollectionManagementReportTXTFileHeader", "TblGeneral") Then
        MsgBox "Collection Management report uploading FAILED because the file hearder doesn't match with the expectations. Please make sure the Collection Management report file has the correct layout.", vbCritical
        Close #1
        DoCmd.Close , MskImportInsuranceCLs
        Exit Sub
    End If
    Close #1

    CurrentDb.Execute "DELETE Tbl_LOGICollectionsManagementReport.* FROM Tbl_LOGICollectionsManagementReport;" ' delete Tbl_LOGICollectionsManagementReport
    DoCmd.TransferText acImportDelim, , "Tbl_LOGICollectionsManagementReport", StrInputFileName, True, "" ' import new collection management report in .txt file
    CurrentDb.Execute "UPDATE Tbl_Customers SET Tbl_Customers.TotalInsurance = null;" ' put null in all CL insurance in Tbl_Customers
    CurrentDb.Execute "UPDATE Tbl_Customers INNER JOIN Tbl_LOGICollectionsManagementReport ON Tbl_Customers.Customer_code = Tbl_LOGICollectionsManagementReport.[Customer Number] SET Tbl_Customers.TotalInsurance = [Tbl_LOGICollectionsManagementReport].[Trade Ins CL];" 'updates the insurance CL according to the collection management report
    MsgBox "Collection Management report uploading completed.", vbOKOnly
End If
DoCmd.Close , MskImportInsuranceCLs
End Sub
