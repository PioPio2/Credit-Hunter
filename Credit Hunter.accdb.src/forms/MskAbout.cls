Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private miRed As Integer
Private miBlue As Integer
Private miGreen As Integer
Private mbFlipRed As Integer
Private mbFlipBlue As Integer
Private mbFlipGreen As Integer
Private delay As Integer

Private Sub Auto_Logo0_DblClick(Cancel As Integer)
    a = MsgBox("Special thanks to: " & Chr(13) & Chr(13) & "My father who bought the first IT book and now he regrets that" & Chr(13) & "My mother who can't still believe how many hours I can spend in front of a pc. Bless her." & Chr(13) & "Manowar and Iron Maiden for all the support" & Chr(13) & "Microsoft for having created MS Access" & Chr(13) & "Logitech that gave me the chance to create this application", , "A simple software easter egg")
End Sub

Private Sub Form_Close()
    wFlags = SND_MEMORY
    PlayIt = sndPlaySound("", wFlags)
End Sub

Private Sub Form_Load()
    Me.Label4.ForeColor = 0
    delay = 0
    SoundName = GetPathExcelDirectory & "about.wav"
    wFlags = SND_ASYNC Or SND_NODEFAULT
    PlayIt = sndPlaySound(SoundName, wFlags)
End Sub

Private Sub Form_Timer()
' Cycle through each of the RGB color values,
' each with a different periodicity. To keep
' from seeing jarring shifts, alternate directions.
' Also, the step values are prime and small in order
' to keep the total cycle long.

    miRed = miRed + mbFlipRed * 2
    If miRed <= 0 Then
        mbFlipRed = 1
        miRed = 0
    End If
    If miRed > 255 Then
        mbFlipRed = -1
        miRed = 255
    End If

    miGreen = miGreen + mbFlipGreen * 3
    If miGreen <= 0 Then
        mbFlipGreen = 1
        miGreen = 0
    End If
    If miGreen > 255 Then
        mbFlipGreen = -1
        miGreen = 255
    End If

    miBlue = miBlue + mbFlipBlue * 5
    If miBlue <= 0 Then
        mbFlipBlue = 1
        miBlue = 0
    End If
    If miBlue > 255 Then
        mbFlipBlue = -1
        miBlue = 255
    End If

    Me.Label4.ForeColor = RGB(miRed, miGreen, miBlue)
    delay = delay + 1
    If delay = 3 Then
        delay = 0
        Me.Label4.Caption = Mid(Me.Label4.Caption, 2) & Left(Me.Label4.Caption, 1)
    End If
End Sub
