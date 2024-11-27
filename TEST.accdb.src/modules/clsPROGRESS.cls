VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPROGRESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' ================================================================================
' CLASSE PER REALIZZARE UNA PROGRESS_BAR XP_STYLE
' AUTORE    :   Alessandro Baraldi
' DATA      :   30/09/2007
' SITO      :   http://www.alessandrobaraldi.it/
' ================================================================================

' ================================================================================
' L'utilizzo di questo codice è gratuito, si richiede di non eliminare
' la paternità dell'autore.
'
' Ringrazio MIRKO(THE TRUSTER) per un suggerimento importante applicato
' nella gestione della PB, al fine di risolvere un noioso problema di FLICKER grafico.
' ================================================================================
'
'   Utilizza una WINDOW'S FORM disegnata RUNTIME, ed associata all'hWnd della
'   sezione in cui risiede il CONTROL_BOX(un Rettangolo) dal quale prende le dimensioni
'   Su questa FORM che possiede, contrariamente alle ACCESS'S FORM, un hDC(Device Context)
'   gli disegno le forme per realizzare la Progressione e quanto serve per lo scopo.
'   Purtroppo per chi non ha familiarità per le API e per la GRAFICA non sarà un esempio
'   semplice da comprendere, quindi ho cercato di realizzare un'interfaccia di utilizzo
'   il più comoda possibile. Come vedrete bastano 3 righe di codice.
'   Per chi volesse addentrarsi nel merito tecnico il codice è molto commentato e strutturato
'   in modo da essere discretamente comprensibile.

' ================================================================================
' I PASSI CHE SERVONO:
' Dopo aver copiato il MODULO DI CLASSE nel Vostro applicativo:
'   ------------------------------------------------------------------------------------------
'   1)  INSERIRE un controllo RETTANGOLO nella FORM e definirne il Bordo TRASPARENTE
'   ------------------------------------------------------------------------------------------
'   2)  Inserire nel modulo di Classe della FORM in cima:
'       Private mPB_O As clsPROGRESS
'   ------------------------------------------------------------------------------------------
'   3)  Su EVENTO LOAD della FORM istanziare le Classi
'       Se si vuole visualizzare l'oggetto subito serve farlo solo dopo che
'       la FORM è visibile, quindi usiamo un TRUCCO, definendo un TIMERINTERVAL=50
'       Altrimenti basterà CREARE la PB al momento dell'utilizzo.
'   ------------------------------------------------------------------------------------------
'   4)  Nell'EVENTO TIMER della FORM creiamo la BARRA associandola al BOX(Rettangolo)
'       mPB_O.CreaPB Me.shPB_O, Orizzontale
'   ------------------------------------------------------------------------------------------
'   5)  Prima di iniziare a POPOLARLA la dobbiamo definire, basta un'istruzione
'       Se ci accontentiamo dei parametri di DEFAULT è sufficiente questo:
'       mPB_O.DefinisciPB
'   ------------------------------------------------------------------------------------------
'   6)  Iniziamo la PROGRESSIONE ed ad ogni nuovo valore lo passeremo al METODO [Progressione]
'       mPB_O.Progressione (x)
'   ------------------------------------------------------------------------------------------
'   7)  Al termine della progressione(100%) la Barra ovviamente rimarrà completamente disegnata
'       quindi se la vogliamo resettare passiamo il valore (0)
'       mPB_O.Progressione(0)
'   ------------------------------------------------------------------------------------------
'   8)  Su EVENTO UNLOAD assicuriamoci di liberare le risorse, in questo caso più che in altri.
'       Set mPB_O = Nothing
' ================================================================================

Public Enum Direzione
    Orizzontale = 0
    Verticale = 1
End Enum

Public Enum ProgressType
    Continua = 1
    Graduata = 2
End Enum

Public Enum ProgressEffect
    Standard = 1
    Gradiente_3D = 2
    Tube_3D = 3
End Enum

Public Event Progress(value As Variant)

Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare PtrSafe Function SetROP2 Lib "GDI32.dll" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare PtrSafe Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare PtrSafe Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare PtrSafe Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare PtrSafe Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare PtrSafe Function PrivateSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

' ================================================================================
' COSTANTI PER LA CONVERSIONE TWIPS/PIXELS
' ================================================================================
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Const DIRECTION_VERTICAL = 1
Private Const DIRECTION_HORIZONTAL = 0

' ================================================================================
' COSTANTI PER LA DEFINIZIONE DELLO STYLE DELLA WINDOW'S-FORM
' ================================================================================
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const SS_BITMAP = &HE&
Private Const WS_EX_WINDOWEDGE = &H100&

' ================================================================================
' COSTANTI PER IL DRAW-MODE(SETROP2/GETROP2)
' ================================================================================
Private Const R2_BLACK As Long = &H1
Private Const R2_NOTMERGEPEN As Long = &H2
Private Const R2_MASKNOTPEN As Long = &H3
Private Const R2_NOTCOPYPEN As Long = &H4
Private Const R2_MASKPENNOT As Long = &H5
Private Const R2_NOT As Long = &H6
Private Const R2_XORPEN As Long = &H7
Private Const R2_NOTMASKPEN As Long = &H8
Private Const R2_MASKPEN As Long = &H9
Private Const R2_NOTXORPEN As Long = &HA
Private Const R2_NOP As Long = &HB
Private Const R2_MERGENOTPEN As Long = &HC
Private Const R2_COPYPEN As Long = &HD
Private Const R2_MERGEPENNOT As Long = &HE
Private Const R2_MERGEPEN As Long = &HF
Private Const R2_WHITE As Long = &H10

' ================================================================================
' COSTANTI PER LO STYLE DELLA PEN
' ================================================================================
Const PS_SOLID = 0


' ================================================================================
' Variabili di tipo Complesso per API
' ================================================================================
Private Type COORDS
    x   As Long
    Y   As Long
    CX  As Long
    CY  As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type POINTAPI
    x                   As Long
    Y                   As Long
End Type

Private Type RGBTRIPLE
    rgbtBlue                    As Byte
    rgbtGreen                   As Byte
    rgbtRed                     As Byte
End Type

Private Type TRIVERTEX
    x                   As Long
    Y                   As Long
    Red                 As Integer 'Ushort value
    Green               As Integer 'Ushort value
    Blue                As Integer 'ushort value
    alpha               As Integer 'ushort
End Type

Private Type GRADIENT_RECT
    UpperLeft           As Long 'In reality this is a UNSIGNED Long
    LowerRight          As Long 'In reality this is a UNSIGNED Long
End Type

' ================================================================================
' VARIABILI INTERNE PER LA DEFINIZIONE GRAFICA DELLA WINDOW'S-FORM
' ================================================================================
Private WP_hWnd                 As Long             ' Handle della WINDOW
Private WP_hDC                  As Long             ' ContextDevice della WINDOW
Private WP_mCOORD               As COORDS           ' COORDINATE della WINDOW RELATIVE ALLA SECIONT
Private WP_RECT                 As RECT             ' AREA DELLA WINDOW
Private WP_AREA                 As RECT             ' AREA GRAFICA DEDICATA ALLA PB

' ================================================================================
' VARIABILI INTERNE PER LA DEFINIZIONE DELLA FORM OSPITANTE
' ================================================================================
Private mFrm                    As Access.Form      ' FORM PARENT
Private mhWndSection            As Long             ' Handle della SECTION
Private mObjContainer           As Access.control   ' Oggetto SHAPE di CONTORNO

' ================================================================================
' COSTANTI GRAFICHE INTERNE PER LA DEFINIZIONE DELLA PROGRESS_BAR
' ================================================================================
Private Const CPB_CornerRND = 1                     ' Raggio di CURVATURA REOUND RECT
Private Const CPB_BlockDX = 3                       ' Distanza BLOCKS GRADUATI

' ================================================================================
' VARIABILI INTERNE PER LA DEFINIZIONE DELLA PROGRESS_BAR
' ================================================================================
Private PB_BackColorINI         As Long             ' Colore Reale START Sfumatura
Private PB_BackColorEND         As Long             ' Colore Reale STOP Sfumatura
Private PB_Orientation          As Direzione        ' Orientamento reale sfumatura
Private PB_Type                 As ProgressType     ' Tipo di PB(Continua/Graduata)
Private PB_Effect               As ProgressEffect   ' Tipo di Effetto applicato
Private PB_NumBlock             As Byte             ' Nel caso di Graduata è il Numero dei BLOCKS

Private PB_MIN                  As Long             ' VALORE MINIMO
Private PB_MAX                  As Long             ' VALORE MASSIMO
Private PB_VALUE                As Variant          ' VALORE ISTANTANEO

Private PB_dx                   As Single           ' LARGHEZZA DEI BLOCKS PROPORZIONATA A D(MAX-MIN)
Private PB_dxRECT               As Single           ' LARGHEZZA DEI BLOCKS PROPORZIONATA A WP_AREA(RIGHT-LEFT)

' ================================================================================
' VARIABILI INTERNE PER LA GESTIONE DELLA PROGRESS_BAR
' ================================================================================
Private OldProgrValue           As Long             ' Valore precedente per la discriminazione del quanto(GRADUATO)
Private RGB_CORRENTE            As RGBTRIPLE        ' COLORE ATTUALE
Private RGB_PRECEDENTE          As RGBTRIPLE        ' COLORE DEL PRECEDENTE INCREMENTO
Private RGB_FINALE              As RGBTRIPLE        ' COLORE FINALE
Private RGB_INIZIALE            As RGBTRIPLE        ' COLORE INIZIALE
Private DiffRosso               As Integer          ' D[rosso](ColoreFinale-ColoreIniziale)
Private DiffVerde               As Integer          ' D[Verde](ColoreFinale-ColoreIniziale)
Private DiffBlu                 As Integer          ' D[Blue](ColoreFinale-ColoreIniziale)


Property Get hwnd() As Long
    hwnd = WP_hWnd
End Property

Property Get hdc() As Long
    hdc = WP_hDC
End Property

Property Get hWndSection() As Long
    hWndSection = mhWndSection
End Property

Property Get FormParent() As Access.Form
    Set FormParent = mFrm
End Property

Property Get Contenitore() As Access.control
    Set Contenitore = mObjContainer
End Property

Property Let ColoreIniziale(value As Long)
    PB_BackColorINI = value
End Property
Property Get ColoreIniziale() As Long
    ColoreIniziale = PB_BackColorINI
End Property

Property Let ColoreFinale(value As Long)
    PB_BackColorEND = value
End Property
Property Get ColoreFinale() As Long
    ColoreFinale = PB_BackColorEND
End Property

Property Let Orientamento(value As Direzione)
    PB_Orientation = value
End Property
Property Get Orientamento() As Direzione
    Orientamento = PB_Orientation
End Property

Property Let NumeroElementi(value As Byte)
    PB_NumBlock = value
End Property
Property Get NumeroElementi() As Byte
    NumeroElementi = PB_NumBlock
End Property

Property Let TipoProgressione(value As ProgressType)
    PB_Type = value
End Property
Property Get TipoProgressione() As ProgressType
    TipoProgressione = PB_Type
End Property

Property Let EffettoProgressione(value As ProgressEffect)
    PB_Effect = value
End Property
Property Get EffettoProgressione() As ProgressEffect
    EffettoProgressione = PB_Effect
End Property

Property Let Minimo(value As Long)
    PB_MIN = value
End Property
Property Get Minimo() As Long
    Minimo = PB_MIN
End Property

Property Let Massimo(value As Long)
    PB_MAX = value
End Property
Property Get Massimo() As Long
    Massimo = PB_MAX
End Property

Property Get value() As Long
    value = PB_VALUE
End Property

Public Sub DefinisciPB(Optional mTipoProgressione As ProgressType = 2, _
                       Optional mEffettoProgressione As ProgressEffect = 3, _
                       Optional mMinimo As Long = 0, _
                       Optional mMassimo As Long = 100, _
                       Optional mColoreIniziale As Long = vbWhite, _
                       Optional mColoreFinale As Long = vbRed, _
                       Optional mNumeroElementi As Byte = 18)

    Me.TipoProgressione = mTipoProgressione
    Me.EffettoProgressione = mEffettoProgressione
    Me.Minimo = mMinimo
    Me.Massimo = mMassimo
    Me.ColoreIniziale = mColoreIniziale
    Me.ColoreFinale = mColoreFinale
    Me.NumeroElementi = mNumeroElementi

    ' ================================================================================
    'Scompongo il colore iniziale nelle tre componenti RGB
    ' ================================================================================
    RGB_INIZIALE.rgbtRed = CByte(PB_BackColorINI And 255)
    RGB_INIZIALE.rgbtGreen = CByte((PB_BackColorINI And 65280) / &H100)
    RGB_INIZIALE.rgbtBlue = CByte((PB_BackColorINI And 16711680) / &H10000)
    ' ================================================================================
    'Idem per il colore finale
    ' ================================================================================
    RGB_FINALE.rgbtRed = CByte(PB_BackColorEND And 255)
    RGB_FINALE.rgbtGreen = CByte((PB_BackColorEND And 65280) / &H100)
    RGB_FINALE.rgbtBlue = CByte((PB_BackColorEND And 16711680) / &H10000)
    ' ================================================================================
    'Mi calcolo la differenza tra le tre componenti dei due colori
    ' ================================================================================
    DiffBlu = CInt(RGB_FINALE.rgbtBlue) - CInt(RGB_INIZIALE.rgbtBlue)
    DiffVerde = CInt(RGB_FINALE.rgbtGreen) - CInt(RGB_INIZIALE.rgbtGreen)
    DiffRosso = CInt(RGB_FINALE.rgbtRed) - CInt(RGB_INIZIALE.rgbtRed)

    ' ================================================================================
    ' Arrotondo per eccesso al fine di essere certo che arrivi fino al 100%
    ' del limite grafico, l'eccesso lo gestisco dopo limitandolo.
    ' ================================================================================
    PB_dx = Arrotonda_Eccesso((PB_MAX - PB_MIN) / PB_NumBlock, 0)

    If Me.Orientamento = Orizzontale Then
        ' ================================================================================
        ' DEFINISCO LA DIMENSIONE DEI BLOCKS ORIZONATALI
        ' ================================================================================
        PB_dxRECT = Arrotonda_Eccesso((WP_AREA.Right - WP_AREA.Left) / PB_NumBlock, 0)
    Else
        ' ================================================================================
        ' DEFINISCO LA DIMENSIONE DEI BLOCKS VERTICALI
        ' ================================================================================
        PB_dxRECT = Arrotonda_Eccesso((WP_AREA.Bottom - WP_AREA.Top) / PB_NumBlock, 0)
    End If

End Sub

Public Function CreaPB(mObj As Access.control, Optional Verso As Direzione = Orizzontale) As Boolean
    ' =========================================================================
    ' Stili della Window API
    ' =========================================================================
    Dim lngExStyle              As Long
    Dim lngStyle                As Long

    CreaPB = False

    If mObj Is Nothing Then Exit Function
    Me.Orientamento = Verso
    Set mObjContainer = mObj
    ' =========================================================================
    ' RICAVO LA FORM PARENT DEL CONTAINER
    ' =========================================================================
    Set mFrm = GetForm(mObj)
    ' =========================================================================
    ' RICAVO L'hWnd DELLA SEZIONE IN CUI SI TROVA IL CONTAINER
    ' =========================================================================
    mhWndSection = fFindSectionhWnd(mObj)
    If mhWndSection = 0 Then Exit Function

    With WP_mCOORD
        .x = fTwipsToPixels(mObjContainer.Left, 0)
        .Y = fTwipsToPixels(mObjContainer.Top, 1)
        .CX = fTwipsToPixels(mObjContainer.Width, 0)
        .CY = fTwipsToPixels(mObjContainer.Height, 1)
    End With

    ' =========================================================================
    ' LIBERO LA MEMORIA DA PRECEDENTI ERRONEI TENTATIVI
    ' =========================================================================
    Call ReleaseDC(WP_hWnd, WP_hDC)
    Call DestroyWindow(WP_hWnd)

    lngStyle = WS_POPUP Or WS_VISIBLE Or SS_BITMAP
    lngExStyle = WS_EX_WINDOWEDGE

    ' =========================================================================
    ' CREO LA WINDOW_FORM STATIC
    ' =========================================================================
    WP_hWnd = CreateWindowEx(lngExStyle, _
                            "STATIC", _
                            "", _
                            lngStyle, _
                            WP_mCOORD.x, WP_mCOORD.Y, WP_mCOORD.CX, WP_mCOORD.CY, _
                            mFrm.hwnd, _
                            0, _
                            Application.hWndAccessApp, 0)

    If WP_hWnd = 0 Then Exit Function
    ' ================================================================================
    ' Associo la WINDOW all'hWnd della SECTION.
    ' ================================================================================
    Call SetParent(WP_hWnd, mhWndSection)

    ' ================================================================================
    ' Ricavo il Context Device
    ' ================================================================================
    WP_hDC = GetDC(WP_hWnd)
    If WP_hDC = 0 Then Exit Function

    ' ================================================================================
    ' Ricavo le dimensioni della WINDOW
    ' ================================================================================
    Call GetClientRect(WP_hWnd, WP_RECT)

    ' ================================================================================
    ' Detraggo 2 all'area RECT per definire l'AREA di disegno reale
    ' ================================================================================
    With WP_AREA
        .Left = WP_RECT.Left + 2
        .Top = WP_RECT.Top + 2
        .Right = WP_RECT.Right - 2
        .Bottom = WP_RECT.Bottom - 2
    End With
    ' ================================================================================
    ' Coloro lo sfondo del colore della Sezione Ospitante
    ' ================================================================================
    Call DrawBackGround
    CreaPB = True
End Function

Public Function Progressione(ByVal value As Single)

    Dim mX1                     As Long     ' COORDINATA X
    Dim mY1                     As Long     ' COORDINATA Y
    Dim mX2                     As Long     ' COORDINATA CX
    Dim mY2                     As Long     ' COORDINATA CY
    Dim FillGradient            As Direzione

    ' ================================================================================
    ' VARIABILI PER BARRA A GRADIENTE
    ' ================================================================================
    Dim x                       As Long     ' VARIABILE PER LA PROGRESS GRADIENTE
    Dim XFINE                   As Long

    If value > PB_MAX Or value < PB_MIN Then Exit Function

    ' =========================================================================
    ' Con gradiente 3D non la direzione della Sfumatura è opposta
    ' alla direzione della Barra.
    ' Con barra Orizzontale sfumatura Verticale
    ' =========================================================================
    If PB_Effect = Gradiente_3D Then
        FillGradient = Abs(PB_Orientation - 1)
    Else
        FillGradient = PB_Orientation
    End If

    ' ================================================================================
    If PB_Type = Continua Then    ' PROGRESS BAR CONTINUA
    ' ================================================================================

            If OldProgrValue > value Then
                    ' ================================================================================
                    ' Se OldProgrValue>Value significa che in NUOVO(Value) è più basso.
                    ' VALUE---->DIMINUISCE (CANCELLO L'ECCEDENZA)
                    ' Ovviamente terrò conto della DIREZIONE(Orizzontale/Verticale)
                    ' ================================================================================
                    If PB_Orientation = Orizzontale Then
                            mX1 = SCALATURA(value, Orizzontale)
                            mY1 = WP_AREA.Top

                    Else
                            mX1 = WP_AREA.Left
                            mY1 = SCALATURA(value, Verticale)
                    End If

                    mX2 = WP_AREA.Right
                    mY2 = WP_AREA.Bottom

                    Call DrawSfondo(WP_hDC, mX1, mY1, mX2, mY2)
                    OldProgrValue = value

            Else
                    ' ================================================================================
                    ' Se OldProgrValue<Value significa che in NUOVO(Value) è più alto.
                    ' VALUE---->AUMENTA (DISEGNO L'INCREMENTO)
                    ' Ovviamente terrò conto della DIREZIONE(Orizzontale/Verticale)
                    ' ================================================================================
                    mX1 = WP_AREA.Left
                    mY1 = WP_AREA.Top

                    If PB_Orientation = Orizzontale Then
                            mX2 = SCALATURA(value, Orizzontale)
                            mY2 = WP_AREA.Bottom

                    Else
                            mX2 = WP_AREA.Right
                            mY2 = SCALATURA(value, Verticale)
                    End If

                    ' =========================================================================
                    ' Ora devo tenere conto dell'EFFETTO
                    ' STANDARD      :la sfumatura è progressiva da SX a DX, quindi devo ricalcolare
                    '               il GRADIENTE dal punto in cui riprendo il disegno
                    ' GRADIENTE_3D  :la sfumatura è opposta al verso di progressione, quindi sarà
                    '               sempre uguale.
                    ' TUBE_3D       :Come il GRADIENTE_3D ma inverto i colori per l'effetto TUBE
                    ' =========================================================================
                    If PB_Effect = Standard Then
                            ' =========================================================================
                            ' Calcolo il valore corrente delle tre componenti
                            ' =========================================================================
                            RGB_CORRENTE = RGB_RICALC(RGB_INIZIALE, value)
                            Call DrawGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_INIZIALE, RGB_CORRENTE, CPB_CornerRND, FillGradient)

                    ElseIf PB_Effect = Gradiente_3D Then
                            Call DrawGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_INIZIALE, RGB_FINALE, CPB_CornerRND, FillGradient)

                    Else    ' TUBE_3D
                            Call Draw3DGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_FINALE, RGB_INIZIALE, CPB_CornerRND, FillGradient)
                    End If
                    OldProgrValue = value

            End If
    ' ================================================================================
    Else    ' GRADUATA
    ' ================================================================================
    ' Nel caso di TYPE=GRADUATA i controlli avvengono per QUANTI, avendo suddiviso
    ' la Barra in BLOCKS.
    ' Se la VARIAZIONE è MAGGIORE/MINORE della QUANTIZZAZIONE allora viene processato
    ' Nel caso sia minore evito di disegnare(ovviamente non aggiorno [OldProgrValue])
    ' ================================================================================

            If value - PB_MIN < PB_dx Then
                    ' ================================================================================
                    ' SE IL NUOVO VALORE < DELTA EQUIVALE A ZERO GRAFICO - Cancello il disegno precedente
                    ' ================================================================================
                    Call DrawSfondo(WP_hDC, WP_AREA.Left, WP_AREA.Top, WP_AREA.Right, WP_AREA.Bottom)
                    OldProgrValue = value

            ElseIf (OldProgrValue \ PB_dx) = (value \ PB_dx) And Not (value = PB_MAX) Then
                    ' ================================================================================
                    ' La variazione è minore del QUANTO quindi non devo disegnare e non aggiorno
                    ' OldProgrValue altrimenti non disegnerei maì andando in inseguimento...!
                    ' ================================================================================

            ElseIf (OldProgrValue \ PB_dx) > (value \ PB_dx) Then
                    ' ================================================================================
                    ' Se QUANTO(OldProgrValue)>QUANTO(Value) significa che in QUANTO(Value) è più basso.
                    ' VALUE---->DIMINUISCE (CANCELLO L'ECCEDENZA)
                    ' Ovviamente terrò conto della DIREZIONE(Orizzontale/Verticale)
                    ' ================================================================================
                    If PB_Orientation = Orizzontale Then
                            mX1 = (value \ PB_dx) * PB_dxRECT
                            mY1 = WP_AREA.Top
                    Else
                            mX1 = WP_AREA.Left
                            mY1 = (value \ PB_dx) * PB_dxRECT
                    End If
                    mX2 = WP_AREA.Right
                    mY2 = WP_AREA.Bottom

                    Call DrawSfondo(WP_hDC, mX1, mY1, mX2, mY2)
                    OldProgrValue = value

            ElseIf (OldProgrValue \ PB_dx) < (value \ PB_dx) Or (value = PB_MAX) Then
                    ' ================================================================================
                    ' Se OldProgrValue<Value significa che in NUOVO(Value) è più alto.
                    ' VALUE---->AUMENTA (DISEGNO L'INCREMENTO)
                    ' Ovviamente terrò conto della DIREZIONE(Orizzontale/Verticale)
                    ' In questo caso la variazione potrebbe essere maggiore di (1) Quanto
                    ' quindi per essere sicuro di disegnare tutto faccio un ciclo dal
                    ' QUANTO precedente a quello attuale e li disegno.
                    ' ================================================================================
                    XFINE = Arrotonda_Eccesso((value / PB_dx), 0)
                    For x = (OldProgrValue \ PB_dx) To XFINE

                            If PB_Orientation = Orizzontale Then
                                    mX1 = VALIDA_RECT(x * PB_dxRECT, Orizzontale)
                                    mY1 = WP_AREA.Top
                                    mX2 = VALIDA_RECT(mX1 + PB_dxRECT - CPB_BlockDX, Orizzontale)
                                    mY2 = WP_AREA.Bottom
                            Else
                                    mX1 = WP_AREA.Left
                                    mY1 = VALIDA_RECT(x * PB_dxRECT, Verticale)
                                    mX2 = WP_AREA.Right
                                    mY2 = VALIDA_RECT(mY1 + PB_dxRECT - CPB_BlockDX, Verticale)
                            End If

                            ' =========================================================================
                            ' Ora devo tenere conto dell'EFFETTO
                            ' STANDARD      :la sfumatura è progressiva da SX a DX, quindi devo ricalcolare
                            '               il GRADIENTE dal punto in cui riprendo il disegno
                            ' GRADIENTE_3D  :la sfumatura è opposta al verso di progressione, quindi sarà
                            '               sempre uguale.
                            ' TUBE_3D       :Come il GRADIENTE_3D ma inverto i colori per l'effetto TUBE
                            ' =========================================================================

                            If PB_Effect = Standard Then
                                    ' ================================================================================
                                    'Calcolo il valore corrente delle tre componenti
                                    ' ================================================================================
                                    RGB_CORRENTE = RGB_RICALC(RGB_INIZIALE, x * PB_dx)
                                    If x = PB_MIN Then
                                            RGB_PRECEDENTE.rgbtBlue = RGB_INIZIALE.rgbtBlue
                                            RGB_PRECEDENTE.rgbtGreen = RGB_INIZIALE.rgbtGreen
                                            RGB_PRECEDENTE.rgbtRed = RGB_INIZIALE.rgbtRed
                                    End If
                                    Call DrawGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_PRECEDENTE, RGB_CORRENTE, CPB_CornerRND, FillGradient)
                                    RGB_PRECEDENTE = RGB_CORRENTE

                            ElseIf PB_Effect = Gradiente_3D Then
                                    Call DrawGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_INIZIALE, RGB_FINALE, CPB_CornerRND, FillGradient)
                            Else
                                    Call Draw3DGradientRoundRect(WP_hDC, mX1, mY1, mX2, mY2, RGB_FINALE, RGB_INIZIALE, CPB_CornerRND, PB_Orientation)
                            End If

                    Next
                    OldProgrValue = value

            End If
    End If

    PB_VALUE = value
    RaiseEvent Progress(value)
    ' =========================================================================
    ' Restituisco il FOCUS alla FORM_PARENT in quanto lo perde quando
    ' si lavora sulla WINDOW_STATIC
    ' =========================================================================
    Call PrivateSetFocus(mFrm.hwnd)

End Function

Public Function DrawBackGround()
    ' ================================================================================
    ' DISEGNO IL CONTENITORE DELLA PB SFUMATO E BORDATO STILE XP
    ' ================================================================================

    Dim hRM                 As Long         ' DRAWMODE

    ' ================================================================================
    ' Memorizzo il DRAWMODE
    ' ================================================================================
    hRM = GetROP2(WP_hDC)

    ' ================================================================================
    ' Imposto il DRAWMODE=R2_COPYPEN
    ' ================================================================================
    Call SetROP2(WP_hDC, R2_COPYPEN)

    ' ================================================================================
    ' Disegno il BORDO ESTERNO
    ' ================================================================================
    Call DrawRoundRect(WP_hDC, _
                       WP_RECT.Left, _
                       WP_RECT.Top, _
                       WP_RECT.Right - WP_RECT.Left, _
                       WP_RECT.Bottom - WP_RECT.Top, 8421504, 5, 1)


    ' ================================================================================
    ' Disegno lo sfondo sfumato
    ' ================================================================================
    Call DrawSfondo(WP_hDC, _
                    WP_RECT.Left + 2, _
                    WP_RECT.Top + 2, _
                    WP_RECT.Right - WP_RECT.Left - 2, _
                    WP_RECT.Bottom - WP_RECT.Top - 2)

    ' ================================================================================
    ' Imposto il DRAWMODE = MASKPEN
    ' ================================================================================
    Call SetROP2(WP_hDC, R2_MASKPEN)
    ' ================================================================================
    ' Disegno la sfumatura interna per l'effetto 3D
    ' ================================================================================
    Call DrawRoundRect(WP_hDC, _
                       WP_RECT.Left + 1, _
                       WP_RECT.Top + 1, _
                       WP_RECT.Right - WP_RECT.Left - 1, _
                       WP_RECT.Bottom - WP_RECT.Top - 1, 14276560, 3, 1)

    ' ================================================================================
    ' Ripristino il DRAWMODE
    ' ================================================================================
    Call SetROP2(WP_hDC, hRM)
End Function

Private Sub DrawSfondo(hdc As Long, _
                       ByVal x As Long, _
                       ByVal Y As Long, _
                       ByVal CX As Long, _
                       ByVal CY As Long)

    ' ================================================================================
    ' QUESTA ROUTINE DISEGNA LO SFONDO PER CANCELLARE GLI ECCESSI DI PROGRESS_BAR
    ' ================================================================================

    Dim CI                  As RGBTRIPLE    ' COLORE INIZIALE BOX CONTENITORE(GRAFICO)
    Dim CF                  As RGBTRIPLE    ' COLORE FINALE BOX CONTENITORE(GRAFICO)

    Dim hRM                 As Long

    hRM = GetROP2(hdc)

    ' ================================================================================
    ' Imposto il DRAWMODE=R2_COPYPEN
    ' ================================================================================
    Call SetROP2(hdc, R2_COPYPEN)

    ' ================================================================================
    ' Imposto i colori dello sfondo sfumato
    ' ================================================================================
    If Me.Orientamento = Verticale Then
        CI.rgbtBlue = 255
        CI.rgbtGreen = 255
        CI.rgbtRed = 255

        CF.rgbtBlue = 215
        CF.rgbtGreen = 215
        CF.rgbtRed = 215
    Else
        CI.rgbtBlue = 215
        CI.rgbtGreen = 215
        CI.rgbtRed = 215

        CF.rgbtBlue = 255
        CF.rgbtGreen = 255
        CF.rgbtRed = 255
    End If
    ' ================================================================================
    ' Disegno lo sfondo sfumato
    ' ================================================================================
    Call DrawGradientRoundRect(hdc, CX, CY, x, Y, CI, CF, 0, Abs(Me.Orientamento - 1))
    Call SetROP2(hdc, hRM)

End Sub

Private Sub DrawGradientRoundRect(hdc As Long, _
                                  ByVal x As Long, _
                                  ByVal Y As Long, _
                                  ByVal CX As Long, _
                                  ByVal CY As Long, _
                                  RGBINI As RGBTRIPLE, _
                                  RGBEND As RGBTRIPLE, _
                                  ByVal Curvatura As Long, _
                                  ByVal VERS As Direzione)
    ' ================================================================================
    ' Special TANKS to Mike D. Sutton (Microsoft Visual Basic MVP)
    ' ================================================================================
    Dim vertR(1)                As TRIVERTEX
    Dim gRect                   As GRADIENT_RECT

    Dim hClipRgn                As Long
    Dim hPen                    As Long
    Dim hOldPen                 As Long
    ' ================================================================================
    ' Create and selection clipping region for fill
    ' ================================================================================
    hClipRgn = CreateRoundRectRgn(x, Y, CX, CY, Curvatura, Curvatura)
    Call SelectClipRgn(hdc, hClipRgn)
    ' ================================================================================
    ' Definisce il colore e punto INIZIALE
    ' ================================================================================
    With vertR(0)
        .x = x
        .Y = Y
        .Red = MakeInt(0, RGBINI.rgbtRed)
        .Green = MakeInt(0, RGBINI.rgbtGreen)
        .Blue = MakeInt(0, RGBINI.rgbtBlue)
        .alpha = 0&
    End With

    ' ================================================================================
    ' Definisce il colore e punto FINALE
    ' ================================================================================
    With vertR(1)
        .x = CX
        .Y = CY
        .Red = MakeInt(0, RGBEND.rgbtRed)
        .Green = MakeInt(0, RGBEND.rgbtGreen)
        .Blue = MakeInt(0, RGBEND.rgbtBlue)
        .alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GradientFillRect hdc, vertR(0), 2, gRect, 1, VERS

    ' ================================================================================
    ' Deseleziona ed elimina la Clipping Region
    ' ================================================================================
    Call SelectClipRgn(hdc, 0&)
    Call DeleteObject(hClipRgn)

End Sub

Private Sub Draw3DGradientRoundRect(hdc As Long, _
                                    ByVal x As Long, _
                                    ByVal Y As Long, _
                                    ByVal CX As Long, _
                                    ByVal CY As Long, _
                                    RGBINI As RGBTRIPLE, _
                                    RGBEND As RGBTRIPLE, _
                                    ByVal Curvatura As Long, _
                                    ByVal VERS As Direzione)

    ' =========================================================================
    ' Special TANKS to Mike D. Sutton (Microsoft Visual Basic MVP)
    ' =========================================================================
    Dim vertR(1)                As TRIVERTEX
    Dim gRect                   As GRADIENT_RECT

    Dim hClipRgn                As Long
    Dim hPen                    As Long
    Dim hOldPen                 As Long

    Dim FillVERS                As Direzione

    FillVERS = Abs(PB_Orientation - 1)

    ' =========================================================================
    ' Create and selection clipping region for fill
    ' =========================================================================
    hClipRgn = CreateRoundRectRgn(x, Y, CX, CY, Curvatura, Curvatura)
    Call SelectClipRgn(hdc, hClipRgn)

    ' =========================================================================
    ' Draw gradient fill From ColorINIT in the 1st HalfPart
    ' =========================================================================
    With vertR(0)
        .x = x
        .Y = Y
        .Red = MakeInt(0, RGBINI.rgbtRed)
        .Green = MakeInt(0, RGBINI.rgbtGreen)
        .Blue = MakeInt(0, RGBINI.rgbtBlue)
        .alpha = 0&
    End With

    ' =========================================================================
    ' to ColorEND
    ' =========================================================================
    With vertR(1)
        If VERS = Orizzontale Then
            .x = CX
            .Y = CY / 3
        Else
            .x = CX / 3
            .Y = CY
        End If
        .Red = MakeInt(0, RGBEND.rgbtRed)
        .Green = MakeInt(0, RGBEND.rgbtGreen)
        .Blue = MakeInt(0, RGBEND.rgbtBlue)
        .alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GradientFillRect hdc, vertR(0), 2, gRect, 1, FillVERS

    ' =========================================================================
    ' Deselect and destroy clipping region
    ' =========================================================================
    Call SelectClipRgn(hdc, 0&)
    Call DeleteObject(hClipRgn)

    ' =========================================================================
    ' Create and selection clipping region for fill
    ' =========================================================================
    hClipRgn = CreateRoundRectRgn(x, Y, CX, CY, Curvatura, Curvatura)
    Call SelectClipRgn(hdc, hClipRgn)

    ' =========================================================================
    ' Draw gradient fill From ColorINIT in the 2nd HalfPart
    ' =========================================================================
    With vertR(0)
        If VERS = Orizzontale Then
            .x = x
            .Y = CY / 3
        Else
            .x = CX / 3
            .Y = Y
        End If
        .Red = MakeInt(0, RGBEND.rgbtRed)
        .Green = MakeInt(0, RGBEND.rgbtGreen)
        .Blue = MakeInt(0, RGBEND.rgbtBlue)
        .alpha = 0&
    End With

    ' =========================================================================
    'to ColorEND
    ' =========================================================================
    With vertR(1)
        .x = CX
        .Y = CY
        .Red = MakeInt(0, RGBINI.rgbtRed)
        .Green = MakeInt(0, RGBINI.rgbtGreen)
        .Blue = MakeInt(0, RGBINI.rgbtBlue)
        .alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    GradientFillRect hdc, vertR(0), 2, gRect, 1, FillVERS

    ' =========================================================================
    ' Deselect and destroy clipping region
    ' =========================================================================
    Call SelectClipRgn(hdc, 0&)
    Call DeleteObject(hClipRgn)

End Sub

Private Sub DrawRoundRect(hdc As Long, _
                          ByVal x As Long, _
                          ByVal Y As Long, _
                          ByVal CX As Long, _
                          ByVal CY As Long, _
                          Colore As Long, _
                          ByVal Curvatura As Long, _
                          Optional Spessore As Byte = 1)

    Dim hPen                    As Long
    Dim hOldPen                 As Long

    hPen = CreatePen(PS_SOLID, Spessore, Colore)
    hOldPen = SelectObject(hdc, hPen)
    Call RoundRect(WP_hDC, x, Y, CX, CY, Curvatura, Curvatura)
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)
End Sub

Private Function VALIDA_RECT(ByVal value As Long, Verso As Direzione) As Long
    ' =========================================================================
    ' VERIFICA CHE IL BLOCK NON ESCA DALL'AREA RECT
    ' =========================================================================
    If Verso = Orizzontale Then
        value = value + WP_AREA.Left
        If value > WP_AREA.Right Then
            VALIDA_RECT = WP_AREA.Right
        Else
            VALIDA_RECT = Fix(value)
        End If
    Else
        value = value + WP_AREA.Top
        If value > WP_AREA.Bottom Then
            VALIDA_RECT = WP_AREA.Bottom
        Else
            VALIDA_RECT = Fix(value)
        End If
    End If

End Function

Private Function SCALATURA(ByVal value As Long, Verso As Direzione) As Long
    ' =========================================================================
    ' SCALA LE DIMENSIONI DEL VALORE RELATIVE A PB_MIN/PB_MAX
    ' ALLE DIMENSIONI RECT DELL'AREA
    ' =========================================================================
    If Verso = Orizzontale Then
        SCALATURA = WP_AREA.Left + (value - PB_MIN) * (WP_AREA.Right - WP_AREA.Left) / (PB_MAX - PB_MIN)
    Else
        SCALATURA = WP_AREA.Top + (value - PB_MIN) * (WP_AREA.Bottom - WP_AREA.Top) / (PB_MAX - PB_MIN)
    End If
End Function

Private Function RGB_RICALC(RGBIN As RGBTRIPLE, ByVal value As Long) As RGBTRIPLE
    ' =========================================================================
    'Calcolo il valore corrente delle tre componenti
    ' =========================================================================
    Dim PERC    As Single
    PERC = ((value - PB_MIN) / (PB_MAX - PB_MIN))
    If PERC > 1 Then PERC = 1
    RGB_RICALC.rgbtBlue = RGBIN.rgbtBlue + DiffBlu * PERC
    RGB_RICALC.rgbtGreen = RGBIN.rgbtGreen + DiffVerde * PERC
    RGB_RICALC.rgbtRed = RGBIN.rgbtRed + DiffRosso * PERC
End Function

Private Function fFindSectionhWnd(ctl As Access.control) As Long
    ' =========================================================================
    ' Thanks to Sthepen Lebans
    ' =========================================================================

    Dim hWnd_LSB As Long
    Dim hWnd_Temp As Long

    ' Window RECT vars
    Dim rc As RECT
    Dim pt As POINTAPI

    ' Loop Counters
    Dim SectionCounter As Long
    Dim ctr As Long

    ' Which Section contains the Control?
    Select Case ctl.Section
        Case acDetail
            '0
            SectionCounter = 2
        Case acHeader
            '1
            SectionCounter = 1
        Case acFooter
            '2
            SectionCounter = 3
        Case Else
        '  ****   NEED ERROR HANDLING! ****
    End Select

    ' Setup SectionCounter
    ' Form Header, Detail and then Footer
    ctr = 1

    ' Let's get first Child Window of the FORM
    hWnd_LSB = GetWindow(mFrm.hwnd, GW_CHILD)

    ' Let's walk through every sibling window of the Form
    Do
        If fGetClassName(hWnd_LSB) = "OFormSub" Then
        ' First OFormSub is the Form's Header. We want the next next one
        ' which is the detail section
            If ctr = SectionCounter Then
                fFindSectionhWnd = hWnd_LSB
                Exit Function
            End If

             ' Increment our Section Counter
            ctr = ctr + 1

        End If

    ' Let's get the NEXT SIBLING Window
    hWnd_LSB = GetWindow(hWnd_LSB, GW_HWNDNEXT)

    ' Let's Start the process from the Top again
    ' Really just an error check
    Loop While hWnd_LSB <> 0

    ' SORRY - NO SECTION Handle is available
    fFindSectionhWnd = 0
End Function

' From Dev Ashish's Site
' The Access Web
' http://www.mvps.org/access/

'******* Code Start *********
Public Function fGetClassName(hwnd As Long)
    Dim strBuffer As String
    Dim lngLen As Long
    Const MAX_LEN = 255
    strBuffer = Space$(MAX_LEN)
    lngLen = GetClassName(hwnd, strBuffer, MAX_LEN)
    If lngLen > 0 Then fGetClassName = Left$(strBuffer, lngLen)
End Function
'******* Code End *********

Private Function GetForm(obj As Access.control) As Form
    Dim mParObj As Object

    Set mParObj = obj.Parent
    ' =========================================================================
    ' Loop_Back until i find PARENT(Type)=FORM
    ' =========================================================================
    Do Until TypeOf mParObj Is Form
        Set mParObj = mParObj.Parent
        DoEvents
    Loop
    Set GetForm = mParObj

End Function

Private Function fctRound(varNr As Variant, Optional varPl As Integer = 2) As Double
    ' =========================================================================
    ' by Konrad Marfurt + ("" by) Luke Chung + Karl Donaubauer
    ' =========================================================================

    'esce se valore non numerico
    If Not IsNumeric(varNr) Then Exit Function

    fctRound = Fix("" & varNr * (10 ^ varPl) + Sgn(varNr) * 0.5) / (10 ^ varPl)

End Function

Public Function Arrotonda_Eccesso(ByVal Numero As Double, Optional decimali As Integer = 2) As Double
    ' =========================================================================
    ' Funzione di GIORGIO RANCATI
    ' Numero  =Numero in ingresso
    ' Decimali=Decimali desiderati
    ' =========================================================================
    Dim Result As Variant 'Risultato
    Dim Nr As Variant     'Numero da trattare
    Dim Dec As Integer    'Numero decimali

    Dec = decimali
    'Controllo massimo numero di decimali
    If Dec > 10 Then Dec = 10

    'Converto il numero in ingresso nel tipo decimal
    Nr = CDec(Nz(Numero, 0))

    'Tolgo l'eventuale segno negativo
    Nr = Abs(Nr)

    'Arrotondo
    Result = Nr * 10 ^ Dec + 0.99999
    Result = Fix(Result) / 10 ^ Dec

    'Riassegno l'eventuale segno negativo e rendo il risultato
    Arrotonda_Eccesso = Result * Sgn(Nz(Numero, 0))
End Function

Function fTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
    ' =========================================================================
    ' Thanks to Sthepen Lebans
    ' Function to convert Twips to pixels for the current screen resolution
    ' Accepts:
    ' lngTwips - the number of twips to be converted
    ' lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
    ' Returns:
    ' the number of pixels corresponding to the given twips
    ' =========================================================================

    On Error GoTo E_HANDLE
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = ReleaseDC(0, lngDeviceHandle)
    fTwipsToPixels = lngTwips / 1440 * lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_HANDLE:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
    Resume fExit
End Function

Function fPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
    ' =========================================================================
    ' Thanks to Sthepen Lebans
    ' Function to convert pixels to twips for the current screen resolution
    ' Accepts:
    ' lngPixels - the number of pixels to be converted
    ' lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
    ' Returns:
    ' the number of twips corresponding to the given pixels
    ' =========================================================================
    On Error GoTo E_HANDLE
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = ReleaseDC(0, lngDeviceHandle)
    fPixelsToTwips = lngPixels * 1440 / lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_HANDLE:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
    Resume fExit
End Function

Private Function MakeInt(ByVal LoByte As Byte, ByVal hiByte As Byte) As Integer
    MakeInt = LoByte Or LongToInt((CLng(hiByte) * &H100))
End Function

Private Function LongToInt(ByVal lngNumber As Long) As Integer
    ' Convert a long Value to an integer
    lngNumber = lngNumber And &HFFFF&
    If lngNumber > &H7FFF Then
      LongToInt = lngNumber - &H10000
    Else
      LongToInt = lngNumber
    End If
End Function

Private Sub Class_Initialize()
    ' INIZIO
    Me.Orientamento = Orizzontale
    Me.TipoProgressione = Continua
    Me.EffettoProgressione = Standard
    Me.Minimo = 0
    Me.Massimo = 100
    Me.ColoreIniziale = vbBlue
    Me.ColoreFinale = vbBlue
    Me.NumeroElementi = 18
End Sub

Private Sub Class_Terminate()
    Call ReleaseDC(WP_hWnd, WP_hDC)
    Call DestroyWindow(WP_hWnd)
End Sub
