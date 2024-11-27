Attribute VB_Name = "basPopupRuntime"
Option Compare Database
Option Explicit

'//   QUESTO MODULO PERMETTE DI GENERARE UN
'//   MENU' POPUP RUNTIME DI TIPO TEMPORANEO
'//   CHE VIENE DISTRUTTO ALLA CHIUSURA DEL PROGETTO.



Private Type PopupMenuVar
    Testo                     As String
    IDIcona                   As Long
    Funzione                  As String
    NuovoGruppo               As Boolean
End Type

Public mArrMnu()              As PopupMenuVar
Public Const mMnuName         As String = "mPopupRuntime"

Public pret As Long

Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As POINTAPI

Global Mouse As POINTAPI




'-------------------------------------------------------------------------------
'  GESTIONE DEI MENU' POPUP TEMPORANEI DI FORM
Public Function ExecMnu(Idx As Integer, Optional parameter As Variant)
   On Error GoTo err_SelezioneServer
   Dim IM As New IMenuMessage
   Set IM = Screen.ActiveForm
   IM.Message Idx, parameter

   Exit Function

err_SelezioneServer:
   MsgBox "ExecMnu : " & Screen.ActiveForm.Name & vbCrLf & err.Description & " (" & err.number & ")"
End Function
'-------------------------------------------------------------------------------

Public Function ShowMyPopup(Optional x As Variant, Optional Y As Variant) As Long

  CreatePopup
  Application.CommandBars(mMnuName).ShowPopup x, Y
  ShowMyPopup = pret

End Function

' Crea il Menù POPUP TEMPORANEO se non esiste.
' L'Oggetto MENU' viene distrutto alla chiusura del progetto
Private Function CreatePopup()

'   Dim myBar                  As Office.CommandBar

   ' Cerca nell'insieme CommandBars l'Oggetto Menù = "mPopupRuntime"
 '  For Each myBar In Application.CommandBars
  '      If myBar.Name = mMnuName And myBar.Type = msoBarTypePopup Then
              ' Se esiste lo aggiorna
   '           Call AddItemPopup(myBar, False)
    '          Exit Function
     '   End If
'   Next

   ' Se siamo quì il Menù non esiste, quindi lo CREA
 '  Set myBar = CommandBars.Add(Name:=mMnuName, Position:=msoBarPopup, _
  '                             Temporary:=True)

   ' Ora esiste lo aggiorna in base all'IDX che ne determina l'aspetto
'   Call AddItemPopup(myBar, True)

 '  Set myBar = Nothing

End Function


'Private Function AddItemPopup(cmb As Office.CommandBar, Mode As Boolean)

 '  Dim myCtlBar               As Office.CommandBarControl
  ' Dim intCurrControl         As Integer
   'Dim x                      As Integer

   ' Il parametro MODE serve per definire  se l'Oggetto
   ' Commandbar = "mPopupRuntime" esisteva per rimuovere tutti gli ITEMS vecchi.
'   If Not Mode Then
 '        For intCurrControl = 1 To cmb.Controls.Count
  '            cmb.Controls(1).Delete
   '      Next
'   End If

   ' Carica dall'array le impostazioni per generare il nuovo Menù
 '  For x = 0 To UBound(mArrMnu)
   '      Set myCtlBar = cmb.Controls.Add(Type:=msoControlButton)
  '
    '     With myCtlBar
'              .Caption = mArrMnu(x).Testo
 '             .FaceId = mArrMnu(x).IDIcona
  '            .OnAction = mArrMnu(x).Funzione
   '           .BeginGroup = mArrMnu(x).NuovoGruppo
    '     End With
'   Next

 '  Set myCtlBar = Nothing

'End Function
 Function CallMenu(Idx As Integer, ChkBox As Boolean)
   '  Questa Routine IMPOSTA la COSTRUZIONE DEL MENU'

   Dim POS           As POINTAPI
   Dim x             As Integer
   Dim lngret        As Long

   Erase mArrMnu

   Select Case Idx
      Case 1
        If ChkBox = False Then
            ReDim mArrMnu(0 To 0)
        Else

         '  MENU' ASSOCIATO ALLA LABEL N° 1
            ReDim mArrMnu(0 To 2)
        End If
         '-------------------------------------------------------------------------
         mArrMnu(0).Testo = "Add attachment(s)"
         mArrMnu(0).IDIcona = 3156
         mArrMnu(0).Funzione = "=ExecMnu(1,1)"
         mArrMnu(0).NuovoGruppo = False
         If ChkBox = True Then
            '-------------------------------------------------------------------------
            mArrMnu(1).Testo = "Remove attachment(s)"
            mArrMnu(1).IDIcona = 2137
            mArrMnu(1).Funzione = "=ExecMnu(1,2)"
            mArrMnu(1).NuovoGruppo = False
            '-------------------------------------------------------------------------
            mArrMnu(2).Testo = "View attachment(s)"
            mArrMnu(2).IDIcona = 2308
            mArrMnu(2).Funzione = "=ExecMnu(1,3)"
            mArrMnu(2).NuovoGruppo = False
         End If
    End Select
    POS = GetCursorPos(Mouse)
    lngret = ShowMyPopup(Mouse.x, Mouse.Y)
End Function
