VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IOutlook

Private OL As Outlook.Application
Private olMail As Outlook.MailItem
Private oAccount As Outlook.Account
Private NAttachment As Long
Private SendImmediately As Boolean

Public Property Get IOutlook_NAttachment() As Integer
    IOutlook_NAttachment = NAttachment
End Property

Public Function IOutlook_CreateOutlook() As Boolean ' Create Outlook application if necessary
    If OL Is Nothing Then
        Set OL = New Outlook.Application
    End If
    IOutlook_CreateOutlook = True
End Function

Public Sub IOutlook_DestroyOutlook()
    Set OL = Nothing
End Sub

Public Function IOutlook_SendEmailFromTemplate(aSubject As String, aHeader As clsCustomerHeader, aDetails As clsCustomerDetails, Attachment As Collection, WordTemplateFileName As String, SendEmail As Boolean) As Boolean
    Call IOutlook_CreateOutlook
    Set olMail = OL.CreateItem(olMailItem)       'create mail item


    Dim User As clsUserS
    Set User = New clsUserS
    Dim oSender As String
    oSender = User.CurrentUserEmailAddress
    Dim oAccount As Outlook.Account              'select email account based on sender email address
    Set oAccount = OL.Session.Accounts.item(oSender)

    olMail.Display False
    olMail.To = aHeader.getEmail
    olMail.Subject = aSubject

    Dim WordObj As clsWord                       'create Word Class
    Set WordObj = New clsWord
    WordObj.IWord_CreateWord
    Call WordObj.IWord_PrepareEmail(GeneralWD, WordTemplateFileName, aHeader, aDetails) 'Open Word template and fill it in

    On Error GoTo CloseDocument

    Dim Editor As Variant
    Set Editor = olMail.GetInspector.WordEditor  ' prepare email editor
    WordObj.CopyContent                          ' copy content of the filled in Word template
    Editor.Content.Paste                         ' paste content from filled in Word template to email

    olMail.SendUsingAccount = oAccount           ' send email as sender email address

    NAttachment = 0
    If Not (Attachment Is Nothing) Then          ' add attachemnt if the case
        If Attachment.Count > 0 Then
            Dim item As Variant
            For Each item In Attachment
                olMail.Attachments.Add item
                NAttachment = NAttachment + 1    ' update n# attachemnt
            Next
        End If
    End If

    If SendEmail = True Then
        olMail.Send                              ' send the email immediately
        SendImmediately = True
    Else
        olMail.Save                              ' save the email in draft
        olMail.Close olSave
        SendImmediately = False
    End If

    WordObj.CloseDocument

    Set oAccount = Nothing                       ' clean things up
    Set olMail = Nothing
    IOutlook_SendEmailFromTemplate = True
    On Error GoTo 0
    Exit Function

CloseDocument:
    WordObj.CloseDocument
    IOutlook_SendEmailFromTemplate = False
End Function

Private Property Get IOutlook_SendImmediately() As Boolean
    IOutlook_SendImmediately = SendImmediately
End Property
