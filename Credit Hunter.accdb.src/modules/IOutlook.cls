VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Property Get SendImmediately() As Boolean

End Property
Public Property Get NAttachment() As Integer

End Property

Public Function SendEmailFromTemplate(aSubject As String, aHeader As clsCustomerHeader, aDetails As clsCustomerDetails, Attachment As Collection, WordTemplateFileName As String, SendEmail As Boolean) As Boolean

End Function
Public Function CreateOutlook() As Boolean

End Function
Public Sub DestroyOutlook()

End Sub
