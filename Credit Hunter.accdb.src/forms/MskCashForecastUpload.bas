Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7591
    DatasheetFontHeight =11
    ItemSuffix =12
    Right =18945
    Bottom =13095
    RecSrcDt = Begin
        0xf875b956f1d6e340
    End
    Caption ="Cash collected"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =6115
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =83
                    Left =1644
                    Top =2891
                    Height =887
                    TabIndex =2
                    Name ="Command4"
                    Caption ="&Select Cash Forecast Excel file"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =83

                    LayoutCachedLeft =1644
                    LayoutCachedTop =2891
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =3778
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =120
                    Top =135
                    Width =7471
                    Height =5980
                    Name ="Box5"
                    LayoutCachedLeft =120
                    LayoutCachedTop =135
                    LayoutCachedWidth =7591
                    LayoutCachedHeight =6115
                End
                Begin TextBox
                    SpecialEffect =1
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2269
                    Top =795
                    Width =1985
                    Height =315
                    Name ="Text0"
                    Format ="Medium Date"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2269
                    LayoutCachedTop =795
                    LayoutCachedWidth =4254
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =285
                            Top =799
                            Width =1035
                            Height =315
                            Name ="Label1"
                            Caption ="From date"
                            LayoutCachedLeft =285
                            LayoutCachedTop =799
                            LayoutCachedWidth =1320
                            LayoutCachedHeight =1114
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =1
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2269
                    Top =1309
                    Width =1985
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    Format ="Medium Date"
                    OnClick ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =2269
                    LayoutCachedTop =1309
                    LayoutCachedWidth =4254
                    LayoutCachedHeight =1624
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =285
                            Top =1305
                            Width =765
                            Height =315
                            Name ="Label3"
                            Caption ="To date"
                            LayoutCachedLeft =285
                            LayoutCachedTop =1305
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =226
                    Top =283
                    Width =4819
                    Height =1704
                    Name ="Box10"
                    LayoutCachedLeft =226
                    LayoutCachedTop =283
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =1987
                End
                Begin Label
                    OverlapFlags =247
                    Left =340
                    Top =170
                    Width =3630
                    Height =285
                    Name ="Label11"
                    Caption ="Payments already uploaded in Access"
                    LayoutCachedLeft =340
                    LayoutCachedTop =170
                    LayoutCachedWidth =3970
                    LayoutCachedHeight =455
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashForecastUpload.cls"
