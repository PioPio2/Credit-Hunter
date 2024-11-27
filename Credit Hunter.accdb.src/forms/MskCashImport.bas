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
    Width =20069
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =28545
    Bottom =13935
    TimerInterval =1
    RecSrcDt = Begin
        0x3bfac1a59caee340
    End
    Caption ="Import cash collected"
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =5952
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =170
                    Top =288
                    Width =18657
                    Height =226
                    Name ="shPB_O2"
                    LayoutCachedLeft =170
                    LayoutCachedTop =288
                    LayoutCachedWidth =18827
                    LayoutCachedHeight =514
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =18992
                    Top =283
                    Width =795
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Etichetta99"
                    Caption ="100%"
                    FontName ="Tahoma"
                    LayoutCachedLeft =18992
                    LayoutCachedTop =283
                    LayoutCachedWidth =19787
                    LayoutCachedHeight =553
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =170
                    Top =735
                    Width =19620
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label0"
                    Caption ="Uploading cash collected..."
                    FontName ="Tahoma"
                    LayoutCachedLeft =170
                    LayoutCachedTop =735
                    LayoutCachedWidth =19790
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =170
                    Top =1125
                    Width =19620
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label1"
                    FontName ="Tahoma"
                    LayoutCachedLeft =170
                    LayoutCachedTop =1125
                    LayoutCachedWidth =19790
                    LayoutCachedHeight =1395
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =170
                    Top =1515
                    Width =19614
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label10"
                    FontName ="Tahoma"
                    LayoutCachedLeft =170
                    LayoutCachedTop =1515
                    LayoutCachedWidth =19784
                    LayoutCachedHeight =1785
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =170
                    Top =1905
                    Width =19674
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label11"
                    FontName ="Tahoma"
                    LayoutCachedLeft =170
                    LayoutCachedTop =1905
                    LayoutCachedWidth =19844
                    LayoutCachedHeight =2175
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =170
                    Top =2295
                    Width =19614
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Label12"
                    FontName ="Tahoma"
                    LayoutCachedLeft =170
                    LayoutCachedTop =2295
                    LayoutCachedWidth =19784
                    LayoutCachedHeight =2565
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =4705
                    Top =3004
                    Width =8430
                    Height =2659
                    BorderColor =0
                    Name ="Tbl_Templates"
                    SourceObject ="Form.MskTemplate"

                    LayoutCachedLeft =4705
                    LayoutCachedTop =3004
                    LayoutCachedWidth =13135
                    LayoutCachedHeight =5663
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashImport.cls"
