Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13833
    DatasheetFontHeight =11
    ItemSuffix =2
    Right =10200
    Bottom =13050
    TimerInterval =1
    RecSrcDt = Begin
        0xeba106229eaee340
    End
    Caption ="Import releases file"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
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
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =108
                    Top =227
                    Width =12921
                    Height =226
                    Name ="shPB_O2"
                    LayoutCachedLeft =108
                    LayoutCachedTop =227
                    LayoutCachedWidth =13029
                    LayoutCachedHeight =453
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =13128
                    Top =228
                    Width =645
                    Height =270
                    FontSize =8
                    FontWeight =600
                    Name ="Etichetta99"
                    Caption ="100%"
                    FontName ="Tahoma"
                    LayoutCachedLeft =13128
                    LayoutCachedTop =228
                    LayoutCachedWidth =13773
                    LayoutCachedHeight =498
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4422
                    Top =1020
                    Width =5736
                    Height =340
                    Name ="Label1"
                    Caption ="Status 1"
                    LayoutCachedLeft =4422
                    LayoutCachedTop =1020
                    LayoutCachedWidth =10158
                    LayoutCachedHeight =1360
                End
            End
        End
    End
End
CodeBehindForm
' See "MskImportLeilaniFile2.cls"
