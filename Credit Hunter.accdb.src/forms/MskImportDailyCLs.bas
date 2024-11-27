Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13250
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =19200
    Bottom =13350
    TimerInterval =1
    RecSrcDt = Begin
        0x3bfac1a59caee340
    End
    Caption ="Import credit limit report"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Section
            Height =5952
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =113
                    Top =288
                    Width =12417
                    Height =226
                    Name ="shPB_O2"
                    LayoutCachedLeft =113
                    LayoutCachedTop =288
                    LayoutCachedWidth =12530
                    LayoutCachedHeight =514
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =12636
                    Top =288
                    Width =495
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Etichetta99"
                    Caption ="100%"
                    FontName ="Tahoma"
                    LayoutCachedLeft =12636
                    LayoutCachedTop =288
                    LayoutCachedWidth =13131
                    LayoutCachedHeight =558
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Top =737
                    Width =13101
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label0"
                    Caption ="Running credit limit report...."
                    FontName ="Tahoma"
                    LayoutCachedLeft =113
                    LayoutCachedTop =737
                    LayoutCachedWidth =13214
                    LayoutCachedHeight =1007
                End
            End
        End
    End
End
CodeBehindForm
' See "MskImportDailyCLs.cls"
