Attribute VB_Name = "basBlob"
'***************************************************************************
'************     Modul mdlBinaries für Access >=97     ********************
'*******    Beliebige Dateien binär in der Datenbank speichern    **********
'****************    01/2004, Sascha Trowitzsch     ************************
'***************************************************************************
 
Option Explicit
Option Compare Database
 
'***************************************************************************
'Funktion 'AddBinFile' fügt der Tabelle tblBinary die Datei sFileName hinzu.
' Falls die Tabelle nicht existiert wird sie neu angelegt.
' Ergebnis der Funktion ist True bei Erfolg
'***************************************************************************
'Function AddBinFile(sFileName As String) As Boolean
'Dim F As Integer
'Dim arrBin() As Byte
'Dim RS As dao.Recordset
 
 '   On Error GoTo Errr
 
    'Fehlertests...
  '  If Not tblBinExists(True) Then err.Raise vbObjectError + 1, "mdlBinary", _
                                    "Binärtabelle konnte nicht erstellt werden!"
   ' If Dir(sFileName) = "" Then err.Raise vbObjectError + 2, "mdlBinary", _
                                    "Datei " & sFileName & "existiert nicht!"
    'Datei einlesen in Byte-Array...
'    F = FreeFile
 '   Open sFileName For Binary As #F
  '  ReDim arrBin(LOF(F))
   ' Get #F, , arrBin()
    'Close #F
 
    'Byte-Array in Tabelle in Binärfeld abspeichern (> .AppendChunk!)
'    Set RS = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
 '   RS.AddNew
  '  RS("FileName") = ExtractFileName(sFileName)
   ' RS("binary").AppendChunk arrBin()
    'RS.Update
'    RS.Close
 '   AddBinFile = True
 
'fExit:
 '   Reset
  '  Erase arrBin
   ' Set RS = Nothing
    'Exit Function
'Errr:
 '   MsgBox err.Description
  '  Resume fExit
'End Function
'*****************************************************************************
'Funktion 'RestoreBinFile' stellt eine Datei aus der Binär-Tabelle wieder her.
' sFileName ist Dateiname (ohne Pfad).
' sPath ist das Verzeichnis, in dem die Datei wiederhergestellt werden soll.
' Overwrite ist optional und standardmäßig True,
'    d.h. eine bereits existierende Datei gleichen Namens wird überschrieben.
' Ergebnis der Funktion ist True bei Erfolg
'*****************************************************************************
Function RestoreBinFile(sFileName, sPath As String, Optional Overwrite As Boolean = True) As Boolean
Dim F As Integer
Dim LSize As Long
Dim arrBin() As Byte
Dim RS As DAO.Recordset
 
    On Error GoTo Errr
 
    If Not tblBinExists Then err.Raise vbObjectError + 3, "mdlBinary", _
                            "Binärtabelle 'tblBinary' existiert nicht in dieser Datenbank!"
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    If Dir(sPath, vbDirectory) = "" Then err.Raise vbObjectError + 4, "mdlBinary", _
                            "Verzeichnis " & sPath & " existiert nicht!"
    If (Dir(sPath & sFileName) <> "") And Not Overwrite Then err.Raise vbObjectError + 4, _
                            "mdlBinary", "Datei " & sFileName & " existiert bereits!"
    Set RS = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    RS.FindFirst "[FileName]='" & sFileName & "'"
    If RS.NoMatch Then
        err.Raise vbObjectError + 5, "mdlBinary", _
                            "Das Binär-File " & sFileName & " existiert nicht in der Tabelle 'tblBinary!'"
    Else
        LSize = RS.Fields("binary").FieldSize
        ReDim arrBin(LSize)
        arrBin = RS.Fields("binary").GetChunk(0, LSize)
        F = FreeFile
        Open sPath & sFileName For Binary As #F
        Put #F, , arrBin
        Close #F
    End If
    RS.Close
    RestoreBinFile = True
 
fExit:
    Reset
    Erase arrBin
    Set RS = Nothing
    Exit Function
Errr:
    MsgBox err.Description
    Resume fExit
End Function
'Hilfsfunktion 'tblBinExists':
'Überprüfen, ob Tabelle "tblBinary" existiert; falls ja, dann Rückgabe: True
'Falls Create=True wird sie erstellt, wenn sie noch nicht existiert
Public Function tblBinExists(Optional Create As Boolean = False) As Boolean
Dim S As String
    On Error Resume Next
    DBEngine(0)(0).TableDefs.Refresh
    S = DBEngine(0)(0).TableDefs("tblBinary").Name
    tblBinExists = (err.number = 0)
    If Create And Not tblBinExists Then tblBinExists = CreateBinTable
End Function
'Hilfsfunktion 'CreateBinTable':
'Erzeugen der Tabelle 'tblBinary'
' Rückgabe: True bei Erfolg
Public Function CreateBinTable() As Boolean
 
    On Error GoTo Errr
 
    DBEngine(0)(0).Execute "CREATE TABLE tblBinary (ID COUNTER CONSTRAINT ID PRIMARY KEY, " & _
                           "FileName CHAR(255) NOT NULL, [binary] IMAGE NOT NULL)"
    'Die Tabelle enthält nun die Felder:
    ' ID (Autowert, pKey)  |   FileName (Text 255)   |     binary (OLE-Feld)
    DBEngine(0)(0).TableDefs.Refresh
    'Der folgende Block bzw. einzelne Elemente ist/sind optional...
    With DBEngine(0)(0).TableDefs("tblBinary")
        .Fields("FileName").Properties.Append .Fields("FileName").CreateProperty( _
                                                    "UnicodeCompression", dbBoolean, True)
        .Properties.Append .CreateProperty("DatasheetFontName", dbText, "Arial")
        .Properties.Append .CreateProperty("DatasheetFontHeight", dbInteger, 8)
        '.Attributes = dbSystemObject  '...Tabelle ist versteckt! '(Nur sichtbar mit Option
                                      ' 'Systemobjekte', kann aber auch dann nicht editiert werden!)
    End With
    CreateBinTable = True
fExit:
    Exit Function
Errr:
    Resume fExit
End Function
'Hilfsfunktion 'ExtractFileName':
'Gibt nur den Dateinamen aus dem vollständige Pfad zurück
Function ExtractFileName(sFilePath As String) As String
Dim N As Long
 
    For N = Len(sFilePath) To 1 Step -1
        If Mid(sFilePath, N, 1) = "\" Then Exit For
    Next N
    ExtractFileName = Mid(sFilePath, N + 1)
 
    'Ab A2000 reicht allein diese Zeile (!):
    'ExtractFileName = Split(sFilePath, "\")(UBound(Split(sFilePath, "\")))
 
End Function


