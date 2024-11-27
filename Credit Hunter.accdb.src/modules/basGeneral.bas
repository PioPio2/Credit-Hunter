Attribute VB_Name = "basGeneral"
Option Compare Database
Option Explicit
Public OL As IOutlook
Public strSelectFile As String
Public label34text As String
Dim FirstFiscalMonthDay As Date

Public GeneralOL As Outlook.Application
Public GeneralWD As Word.Application
Public GeneralExcel As Excel.Application

Public Type CustomerHeader
    Name As String
    TotalOverdue As Currency
    Outstanding As Currency
    CreditLimit As Currency
End Type

Public Sub CheckOutlook()
    If IsNull(OL) Then
        OL = New Outlook.Application
    End If
End Sub


Public Function getIconFromTable(strfilename As String) As Picture

Dim LSize As Long
Dim arrBin() As Byte
Dim RS As DAO.Recordset

    On Error GoTo Errr

    'If Not tblBinExists Then err.Raise vbObjectError + 3, "mdlBinary", _
                            "Binärtabelle 'tblBinary' existiert nicht in dieser Datenbank!"
    Set RS = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    RS.FindFirst "[FileName]='" & strfilename & "'"
    If RS.NoMatch Then
        err.Raise vbObjectError + 5, "mdlBinary", _
                            "Das Binär-File " & strfilename & " existiert nicht in der Tabelle 'tblBinary!'"
    Else
        LSize = RS.Fields("binary").FieldSize - 1
        ReDim arrBin(LSize)
        arrBin = RS.Fields("binary")
'        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    RS.Close

fExit:
    Reset
    Erase arrBin
    Set RS = Nothing
    Exit Function
Errr:
    MsgBox err.Description
    Resume fExit
End Function


Public Function Token(Str As String, Str1 As String, Num As Integer) As String
   ' str = stringa da controllare
   ' str1 = lista dei delimitatori
   ' num = numero del token da prelevare
   ' by Pia Toro & A. Cara
   Dim I As Long
   Dim k As Long
   Dim N As Long
   Dim j As Long
   Dim str0 As String
   Token = ""
   k = 0
   str0 = Left(Str1, 1) & Str & Left(Str1, 1)
   For I = 1 To Len(str0)
      If InStr(Str1, Mid(str0, I, 1)) > 0 Then
         If InStr(Str1, Mid(str0, I + 1, 1)) = 0 Or I = Len(str0) Then
            k = k + 1
               If k = Num + 1 Then
                  Token = Mid(str0, j + 1, N - j)
                  Exit For
               Else
                  j = I
               End If
            End If
      Else
         N = I
      End If
   Next I
End Function


Function QueryCashCollectedSQLParser(Stringa As String, CreditController As Integer, Channel As String) As String
Dim posiz As Integer
    ' "SELECT Tbl_Users.Name, Tbl_Customers.RetailOEM, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Sum(Tbl_CashCollected.Amount) AS [Amount in EUR], Tbl_Cash_Target.CashTargetInEUR FROM (((Tbl_Customers INNER JOIN Tbl_CashCollected ON Tbl_Customers.Customer_code = Tbl_CashCollected.CustomerID) INNER JOIN Tbl_Currencies ON Tbl_CashCollected.Currency = Tbl_Currencies.CurrencyID) INNER JOIN Tbl_Users ON Tbl_Customers.Credit_controller = Tbl_Users.ID) INNER JOIN Tbl_Cash_Target ON Tbl_Users.ID = Tbl_Cash_Target.CControllerID WHERE ((Tbl_Customers.Credit_controller) = " & Combo6.Column(1) & ") GROUP BY Tbl_Users.Name, Tbl_Customers.Name, Tbl_CashCollected.[Payment Date], Tbl_Cash_Target.CashTargetInEUR HAVING (((Tbl_CashCollected.[Payment Date]) >= #" & format(Text0.Value, "mm/dd/yy") & "# And (Tbl_CashCollected.[Payment Date]) <= #" & format(Text2.Value, "mm/dd/yy") & "#)) ORDER BY Sum(Tbl_CashCollected.Amount) DESC;"
    posiz = InStr(1, Stringa, "ORDER BY")
    If CreditController <> 0 Then
        Stringa = Left(Stringa, posiz - 3) & " AND ((Tbl_Users.ID)=" & CreditController & ")" & Mid(Stringa, posiz - 2, 10000)
            'Stringa = Left(Stringa, posiz - 3) & " AND ((Tbl_Users.ID)=1)" & " ORDER BY Sum(Tbl_CashCollected.Amount) DESC;"
    End If
    If Channel <> "" Then
        posiz = InStr(1, Stringa, "HAVING")
        Stringa = Left(Stringa, posiz + 6) & "(Tbl_Customers.RetailOEM='" & Channel & "') AND " & Mid(Stringa, posiz + 7, 10000)
    End If
    QueryCashCollectedSQLParser = Stringa
End Function
