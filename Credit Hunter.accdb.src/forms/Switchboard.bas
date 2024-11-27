Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7740
    DatasheetFontHeight =11
    ItemSuffix =22
    Right =13635
    Bottom =9720
    TimerInterval =500
    Filter ="[ItemNumber] = 0 AND [Argument] = 'Default' "
    RecSrcDt = Begin
        0xc5bd824bbd34e640
    End
    RecordSource ="Switchboard Items"
    Caption ="Main Switchboard"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Aptos"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            ForeThemeColorIndex =2
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderThemeColorIndex =3
            BorderShade =90.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontFamily =0
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Aptos"
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin Section
            Height =5952
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =2355
                    Width =378
                    Height =4770
                    BackColor =8421504
                    BorderColor =0
                    Name ="VerticalShadowBox"
                    LayoutCachedLeft =2355
                    LayoutCachedWidth =2733
                    LayoutCachedHeight =4770
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =1305
                    Width =259
                    Height =259
                    FontSize =10
                    ForeColor =0
                    Name ="Option1"
                    OnClick ="=HandleButtonClick(1)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =1305
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =1564
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =1305
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel1"
                            Caption ="Main Form"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(1)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =1305
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =1545
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =1725
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =1
                    ForeColor =0
                    Name ="Option2"
                    OnClick ="=HandleButtonClick(2)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =1725
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =1984
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =1725
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel2"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(2)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =1725
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =1965
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =2145
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =2
                    ForeColor =0
                    Name ="Option3"
                    OnClick ="=HandleButtonClick(3)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =2145
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =2404
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =2145
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel3"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(3)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =2145
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =2385
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =2565
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =3
                    ForeColor =0
                    Name ="Option4"
                    OnClick ="=HandleButtonClick(4)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =2565
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =2824
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =2565
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel4"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(4)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =2565
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =2805
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =2985
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =4
                    ForeColor =0
                    Name ="Option5"
                    OnClick ="=HandleButtonClick(5)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =2985
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =3244
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =2985
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel5"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(5)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =2985
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =3225
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =3405
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =5
                    ForeColor =0
                    Name ="Option6"
                    OnClick ="=HandleButtonClick(6)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =3405
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =3664
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =3405
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel6"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(6)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =3405
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =3645
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =3825
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =6
                    ForeColor =0
                    Name ="Option7"
                    OnClick ="=HandleButtonClick(7)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =3825
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =4084
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =3825
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel7"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(7)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =3825
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =4065
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3030
                    Top =4245
                    Width =259
                    Height =259
                    FontSize =10
                    TabIndex =7
                    ForeColor =0
                    Name ="Option8"
                    OnClick ="=HandleButtonClick(8)"
                    FontName ="System"

                    LayoutCachedLeft =3030
                    LayoutCachedTop =4245
                    LayoutCachedWidth =3289
                    LayoutCachedHeight =4504
                    ForeThemeColorIndex =-1
                    BackColor =0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =34
                            Left =3390
                            Top =4245
                            Width =3990
                            Height =240
                            FontSize =8
                            BorderColor =0
                            ForeColor =-2147483630
                            Name ="OptionLabel8"
                            FontName ="Tahoma"
                            OnClick ="=HandleButtonClick(8)"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =4245
                            LayoutCachedWidth =7380
                            LayoutCachedHeight =4485
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =223
                    Width =7380
                    Height =660
                    BackColor =8421376
                    BorderColor =8421376
                    Name ="HorizontalHeaderBox"
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =660
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    OverlapFlags =95
                    SpecialEffect =1
                    Left =2685
                    Top =1155
                    Width =4698
                    Name ="HorizontalDividingLine"
                    LayoutCachedLeft =2685
                    LayoutCachedTop =1155
                    LayoutCachedWidth =7383
                    LayoutCachedHeight =1155
                End
                Begin Label
                    OverlapFlags =223
                    TextFontFamily =34
                    Left =2997
                    Top =215
                    Width =4410
                    Height =450
                    FontSize =18
                    BorderColor =0
                    ForeColor =8421504
                    Name ="Label2"
                    Caption ="Credit Hunter"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2997
                    LayoutCachedTop =215
                    LayoutCachedWidth =7407
                    LayoutCachedHeight =665
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =2955
                    Top =170
                    Width =4410
                    Height =450
                    FontSize =18
                    BorderColor =0
                    ForeColor =16777215
                    Name ="Label1"
                    Caption ="Credit Hunter"
                    FontName ="Segoe UI"
                    LayoutCachedLeft =2955
                    LayoutCachedTop =170
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =620
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Image
                    BackStyle =1
                    SizeMode =1
                    Width =2685
                    Height =4770
                    BackColor =8421376
                    BorderColor =0
                    Name ="Picture"

                    LayoutCachedWidth =2685
                    LayoutCachedHeight =4770
                    TabIndex =8
                    BorderThemeColorIndex =-1
                End
            End
        End
    End
End
CodeBehindForm
' See "Switchboard.cls"
