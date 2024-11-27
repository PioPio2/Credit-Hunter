VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private aTime As Single

Public Sub StartTimer()
    aTime = Timer
End Sub

Public Sub StopAndShowTimer(Optional ByVal aMessage As String)
    Debug.Print aMessage & "Seconds: " & Timer - aTime
End Sub
