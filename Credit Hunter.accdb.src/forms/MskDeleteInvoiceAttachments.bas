Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10430
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =13995
    Bottom =13050
    Filter ="DocumentID='337901030/06/2010196100          ' AND CustomerID=41406"
    RecSrcDt = Begin
        0x7903b8a174e1e340
    End
    RecordSource ="Tbl_InvoiceAttachments"
    Caption ="Delete Invoice attachment(s)"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =2848
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =170
                    Top =283
                    Width =10260
                    Height =2565
                    Name ="Tbl_InvoiceAttachments subform"
                    SourceObject ="Form.Tbl_InvoiceAttachments subform"
                    LinkChildFields ="CustomerID;DocumentID"
                    LinkMasterFields ="CustomerID;DocumentID"
                    EventProcPrefix ="Tbl_InvoiceAttachments_subform"

                    LayoutCachedLeft =170
                    LayoutCachedTop =283
                    LayoutCachedWidth =10430
                    LayoutCachedHeight =2848
                End
            End
        End
    End
End
CodeBehindForm
' See "MskDeleteInvoiceAttachments.cls"
