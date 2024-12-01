Attribute VB_Name = "EmailHTML"
Option Compare Database

Sub SendEmailsHTML(strTo As String, strcc As String, strSubject As String, strBody As String, PicturePath As String, ParamArray strFiles())
Const CdoReferenceTypeName = 1
Dim cdomsg, objBP As Variant
Dim Rst As Recordset
Dim x As Integer
If strTo <> "" Then
'    Set rst = New ADODB.Recordset
'    rst.ActiveConnection = CurrentProject.Connection
 '   rst.Open "TblGeneral", , adOpenKeyset, adLockOptimistic, adCmdTable

    Set cdomsg = CreateObject("CDO.message")
    With cdomsg.Configuration.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = Rst.Fields("sendusing") 'NTLM method
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Rst.Fields("smtpserver")
        .item("http://schemas.microsoft.com/cdo/configuration/smptserverport") = Rst.Fields("SMTPserverport")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Rst.Fields("smtpauthenticate")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Rst.Fields("smtpusessl")
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Rst.Fields("smtpconnectiontimeout")
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = DLookup("[Password]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .Update
    End With
    ' build email parts
    With cdomsg
        .To = strTo
        '.to = "alberto_paganini@libero.it"
        If strcc <> "" Then
            .CC = strcc
        End If
       ' .cc = ""
        If DLookup("[EmailSentToSender]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'") Then
            .CC = .CC & "," & DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        End If
        .From = DLookup("[E-mailAddress]", "[Tbl_Users]", "UserName='" & fOSUserName() & "'")
        .Subject = strSubject


        strBody = Replace(strBody, "[picture]", " <img src=""cid:MergedCashTargetReports.bmp"">")
        .HTMLBody = "<body><font face=Arial color=#000080 size=2> " & strBody & "</font></body>"
        '.HTMLBody = .HTMLBody & " <img src=""cid:MergedCashTargetReports.bmp"">"
        '.HTMLBody = .HTMLBody & "</font></body>"

        If PicturePath <> "" Then
            Set objBP = cdomsg.AddRelatedBodyPart(PicturePath, "MergedCashTargetReports.JPG", CdoReferenceTypeName)
            objBP.Fields.item("urn:schemas:mailheader:Content-ID") = "<MergedCashTargetReports.bmp>"
            objBP.Fields.Update
        End If

        'add attachments
'            .Addattachment strFilename
        If UBound(strFiles) > -1 Then
            For x = 0 To UBound(strFiles)
                .Addattachment strFiles(x)
            Next x
        End If
        .Send
    End With
    Rst.Close
    Set Rst = Nothing
    Set cdomsg = Nothing
Else
    x = MsgBox("Main email recipient is missing. The email will not be sent.", vbCritical, "Error")
End If
End Sub
