Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =20749
    DatasheetFontHeight =10
    ItemSuffix =13
    Right =28545
    Bottom =13935
    TimerInterval =1
    RecSrcDt = Begin
        0x82f1979ffb31e340
    End
    Caption ="Upload new statements"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =3344
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =108
                    Top =1092
                    Width =20427
                    Height =325
                    Name ="Etichetta2"
                    Caption =" "
                    LayoutCachedLeft =108
                    LayoutCachedTop =1092
                    LayoutCachedWidth =20535
                    LayoutCachedHeight =1417
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =108
                    Top =1545
                    Width =20427
                    Height =340
                    Name ="Etichetta3"
                    Caption =" "
                    LayoutCachedLeft =108
                    LayoutCachedTop =1545
                    LayoutCachedWidth =20535
                    LayoutCachedHeight =1885
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =90
                    Top =450
                    Width =19470
                    Height =284
                    Name ="shPB_O2"
                    LayoutCachedLeft =90
                    LayoutCachedTop =450
                    LayoutCachedWidth =19560
                    LayoutCachedHeight =734
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =19755
                    Top =450
                    Width =761
                    Height =285
                    FontWeight =700
                    Name ="Etichetta7"
                    Caption ="0%"
                    LayoutCachedLeft =19755
                    LayoutCachedTop =450
                    LayoutCachedWidth =20516
                    LayoutCachedHeight =735
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Top =525
                    Width =1588
                    Height =160
                    Name ="Testo8"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =108
                    Top =2005
                    Width =20427
                    Height =340
                    FontWeight =700
                    Name ="Label10"
                    Caption =" "
                    LayoutCachedLeft =108
                    LayoutCachedTop =2005
                    LayoutCachedWidth =20535
                    LayoutCachedHeight =2345
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =108
                    Top =2465
                    Width =20412
                    Height =340
                    FontWeight =700
                    Name ="Label11"
                    Caption =" "
                    LayoutCachedLeft =108
                    LayoutCachedTop =2465
                    LayoutCachedWidth =20520
                    LayoutCachedHeight =2805
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =108
                    Top =2925
                    Width =20427
                    Height =340
                    FontWeight =700
                    Name ="Label12"
                    Caption =" "
                    LayoutCachedLeft =108
                    LayoutCachedTop =2925
                    LayoutCachedWidth =20535
                    LayoutCachedHeight =3265
                End
            End
        End
    End
End
CodeBehindForm
' See "MskImportStatement2.cls"
