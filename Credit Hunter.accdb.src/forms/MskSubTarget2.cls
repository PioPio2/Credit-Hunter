Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeInsert(Cancel As Integer)
    Me.Text8.value = Me.Parent.Text1.value
    Me.Text10.value = Me.Parent.Text3.value
    Me.Text12.value = Me.Parent.Text5.value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        KeyCode = 0
        Call OpenHelpPage(Me.Name)
    End If
End Sub


Private Sub Text14_LostFocus()
Dim USDExchangeRate  As Currency
Dim EURExchangeRate  As Currency
Dim AmountInUSD As Currency

    USDExchangeRate = DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='USD'")
    EURExchangeRate = DLookup("ExchangeRate", "Tbl_Currencies", "CurrencyID='" & Text6.value & "'")
    '------------------------------
    AmountInUSD = Text14.value * EURExchangeRate / USDExchangeRate
    Text22.value = AmountInUSD
    '--------------------
End Sub

Private Sub Text4_Change()
    Text16.value = Text4.Column(1)
End Sub
