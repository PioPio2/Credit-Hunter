VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWord

Private WordDoc As Word.Document


Public Function IWord_CreateWord() As Boolean
    If GeneralWD Is Nothing Then
        Set GeneralWD = New Word.Application
        Exit Function
    End If
    On Error GoTo ErrorHandler
        Dim Dummy As String
        Dummy = GeneralWD.Name
    On Error GoTo 0
    Exit Function
ErrorHandler:
    Set GeneralWD = New Word.Application
End Function

Private Function IWord_DestroyWord() As Boolean
    Set GeneralWD = Nothing
End Function

Public Function IWord_PrepareEmail(WordTemplateFileName As String, aCustomerHeader As clsCustomerHeader, aCustomerDetails As clsCustomerDetails) As Variant
    Set WordDoc = GeneralWD.Documents.Open(WordTemplateFileName) ' open Word template

    GeneralWD.Visible = False

    Dim Dict As Scripting.Dictionary
    Set Dict = aCustomerHeader.GetHeader
    Dim Key As Variant
    For Each Key In Dict.Keys
        'WDApp.Selection.Find.Text = "<<name>>"
        If Not IsNull(Dict(Key)) Then
            GeneralWD.Selection.Find.Text = "<<" & Key & ">>"
            GeneralWD.Selection.Find.Replacement.Text = Dict(Key)
            GeneralWD.Selection.Find.Execute Replace:=wdReplaceAll
        End If
    Next
    'replace signposts in Word template
    GeneralWD.Selection.Find.Text = "<<name>>"
    GeneralWD.Selection.Find.Replacement.Text = aCustomerHeader.getCustomerName
    GeneralWD.Selection.Find.Execute Replace:=wdReplaceAll

    GeneralWD.Selection.Find.Text = "<<totaloverdue>>"
    GeneralWD.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerDetails.getTotalOverdue, "#,##0.00")
    GeneralWD.Selection.Find.Execute Replace:=wdReplaceAll

    GeneralWD.Selection.Find.Text = "<<outstanding>>"
    GeneralWD.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerDetails.getTotalOutstanding, "#,##0.00")
    GeneralWD.Selection.Find.Execute Replace:=wdReplaceAll

    GeneralWD.Selection.Find.Text = "<<creditlimit>>"
    GeneralWD.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerHeader.getCreditLimit, "#,##0.00")
    GeneralWD.Selection.Find.Execute Replace:=wdReplaceAll

    IWord_PrepareEmail = WordDoc
End Function


Public Function CloseDocument() As Boolean
    WordDoc.Close SaveChanges:=False ' close docuent but don't save
End Function

Public Function CopyContent() As Boolean
    WordDoc.Content.Copy ' copy content of the Word template in clipboard
    CopyContent = True
End Function
