Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public XXX As String


Private Sub News()

Me.lbl_tit.ForeColor = 0


        Dim xmlDoc As Object
        Set xmlDoc = CreateObject("msxml2.domdocument")
        xmlDoc.async = False



            '** LEGGE DA TABELLA
            xmlDoc.Load (Me.lbl_opz.Column(2))
            Set objNodeList = xmlDoc.getElementsByTagName("item")
            Me.lbl_tit = Me.lbl_opz.Column(1)



            '** Il sole 24 ore prima pagina
            Set objNodeList = xmlDoc.getElementsByTagName("item/title")
            'apro recordset
            Dim RS As DAO.Recordset
            Set RS = CurrentDb.OpenRecordset("Tbl_NEWS")
            For I = 0 To (objNodeList.Length - 1)

                '* individua unmis di progressbar
                Dim UnMis As Long
                UnMis = Me.lbl_sfo.Width / objNodeList.Length
                '* incremento progress bar
                Me.lbl_progress.Width = UnMis * (I + 1)

                With RS
                    RS.AddNew
                    RS![TITOLO] = Replace(Replace(Replace(Replace(objNodeList.item(I).XML, "<title>", ""), "</title>", ""), "<![CDATA[", ""), "]]>", "")

                    Set objNodeList2 = xmlDoc.getElementsByTagName("item/link")
                    RS![WEB] = Replace(Replace(Replace(Replace(objNodeList2.item(I).XML, "<link>", ""), "</link>", ""), "<![CDATA[", ""), "]]>", "")

                    Set objNodeList2 = xmlDoc.getElementsByTagName("item/description")
                    RS![Note] = Replace(Replace(Replace(Replace(objNodeList2.item(I).XML, "<description>", ""), "</description>", ""), "<![CDATA[", ""), "]]>", "")

                    RS.Update
                End With
                DoEvents
            Next

End Sub


Private Sub Form_Current()
Me.Img_web.HyperlinkAddress = Me.WEB.value
End Sub

Private Sub Form_Open(Cancel As Integer)
'Me.lbl_progress.Width = 0
'Me.lbl_tit.ForeColor = 255
'Me.lbl_tit = "Connessione alla pagina..."
'DoEvents
''DoCmd.Echo False
'        Me.lbl_opz = 3
        DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE Tbl_NEWS.* FROM Tbl_NEWS;"
        DoCmd.SetWarnings True
'        Call News
        Me.lbl_tit = "Choose news channel ->"
        Me.lbl_progress.Width = 0
        Me.Form.Requery
        Me.lbl_opz.DefaultValue = 1
        Call lbl_opz_AfterUpdate
  '      Me.lbl_opz.SetFocus
   '     Me.lbl_agg.Caption = "Notizie aggiornate al : " & Now()
DoCmd.Echo True
DoEvents
''Me.TimerInterval = 2000
End Sub

Private Sub Form_Timer()
On Error GoTo err

DoCmd.GoToRecord , , acNext
Me.TITOLO.SetFocus
Exit Sub

err:
DoCmd.GoToRecord , , acFirst
Me.TITOLO.SetFocus
End Sub

Private Sub Img_Reconnect_Click()
If Not IsNull(lbl_opz) Then
    Call lbl_opz_AfterUpdate
Else
    MsgBox "Choose the news first", vbCritical, "Attention !!!"
End If
End Sub

Private Sub IntestazioneMaschera_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If XXX <> "corpo" Then
    XXX = "corpo"
    Me.lbl_su.Visible = False
    Me.lbl_giu.Visible = False
Else: End If
End Sub

Private Sub lbl_giu_Click()
On Error GoTo fine:
DoCmd.GoToRecord , , acNext
Me.TITOLO.SetFocus
fine:
End Sub

Private Sub lbl_giu_over_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If XXX <> "lbl_giu_over" Then
    XXX = "lbl_giu_over"
    Me.lbl_su.Visible = False
    Me.lbl_giu.Visible = True
Else: End If
End Sub

Private Sub lbl_opz_AfterUpdate()
Me.lbl_progress.Width = 0
Me.lbl_tit.ForeColor = 255
Me.lbl_tit = "Connecting..."
DoEvents
'DoCmd.Echo False
        Me.lbl_agg.Caption = "...Wait please..."
        Me.lbl_agg.ForeColor = 255
        DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE Tbl_NEWS.* FROM Tbl_NEWS;"
        DoCmd.SetWarnings True
        Call News
        Me.Form.Requery
        Me.lbl_opz.SetFocus
        Me.lbl_agg.ForeColor = 9868950
        Me.lbl_agg.Caption = "News updated as of: " & Now()
        'DoCmd.Echo True
End Sub

Private Sub lbl_su_Click()
On Error GoTo fine:
DoCmd.GoToRecord , , acPrevious
Me.TITOLO.SetFocus
fine:
End Sub

Private Sub lbl_su_over_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If XXX <> "lbl_su_over" Then
    XXX = "lbl_su_over"
    Me.lbl_su.Visible = True
    Me.lbl_giu.Visible = False
Else: End If
End Sub

Private Sub TITOLO_DblClick(Cancel As Integer)
Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Frm_NEWS_Zoom"

    stLinkCriteria = "[ID]=" & Me![ID]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub
