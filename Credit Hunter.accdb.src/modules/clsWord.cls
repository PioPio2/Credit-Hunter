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
    End If
End Function

Public Function IWord_PrepareEmail(WDApp As Word.Application, WordTemplateFileName As String, aCustomerHeader As clsCustomerHeader, aCustomerDetails As clsCustomerDetails) As Variant
    Set WordDoc = WDApp.Documents.Open(WordTemplateFileName) ' open Word template

    WDApp.Visible = False

    Dim Dict As Scripting.Dictionary
    Set Dict = aCustomerHeader.GetHeader
    Dim Key As Variant
    For Each Key In Dict.Keys
        'WDApp.Selection.Find.Text = "<<name>>"
        If Not IsNull(Dict(Key)) Then
        WDApp.Selection.Find.Text = "<<" & Key & ">>"
        WDApp.Selection.Find.Replacement.Text = Dict(Key)
        WDApp.Selection.Find.Execute Replace:=wdReplaceAll
        End If
    Next
    'replace signposts in Word template
    WDApp.Selection.Find.Text = "<<name>>"
    WDApp.Selection.Find.Replacement.Text = aCustomerHeader.getCustomerName
    WDApp.Selection.Find.Execute Replace:=wdReplaceAll

    WDApp.Selection.Find.Text = "<<totaloverdue>>"
    WDApp.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerDetails.getTotalOverdue, "#,##0.00")
    WDApp.Selection.Find.Execute Replace:=wdReplaceAll

    WDApp.Selection.Find.Text = "<<outstanding>>"
    WDApp.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerDetails.getTotalOutstanding, "#,##0.00")
    WDApp.Selection.Find.Execute Replace:=wdReplaceAll

    WDApp.Selection.Find.Text = "<<creditlimit>>"
    WDApp.Selection.Find.Replacement.Text = aCustomerDetails.getCurrency & " " & Format(aCustomerHeader.getCreditLimit, "#,##0.00")
    WDApp.Selection.Find.Execute Replace:=wdReplaceAll

    IWord_PrepareEmail = WordDoc
End Function


Public Function CloseDocument() As Boolean
    WordDoc.Close SaveChanges:=False ' close docuent but don't save
End Function

Public Function CopyContent() As Boolean
    WordDoc.Content.Copy ' copy content of the Word template in clipboard
    CopyContent = True
End Function
