Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12925
    DatasheetFontHeight =11
    ItemSuffix =2
    Right =16560
    Bottom =11865
    RecSrcDt = Begin
        0x23714199dd46e640
    End
    Caption ="Multiple Statements"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =8560
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =453
                    Top =680
                    Width =5220
                    Height =3540
                    Name ="Tbl_CustomersList subform"
                    SourceObject ="Form.Tbl_CustomersList subform"
                    EventProcPrefix ="Tbl_CustomersList_subform"

                    LayoutCachedLeft =453
                    LayoutCachedTop =680
                    LayoutCachedWidth =5673
                    LayoutCachedHeight =4220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =453
                            Top =283
                            Width =2595
                            Height =315
                            Name ="Tbl_CustomersList subform Label"
                            Caption ="Customers"
                            EventProcPrefix ="Tbl_CustomersList_subform_Label"
                            LayoutCachedLeft =453
                            LayoutCachedTop =283
                            LayoutCachedWidth =3048
                            LayoutCachedHeight =598
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6746
                    Top =510
                    Width =2267
                    Height =1417
                    TabIndex =1
                    Name ="Command1"
                    Caption ="Open Excel File"

                    LayoutCachedLeft =6746
                    LayoutCachedTop =510
                    LayoutCachedWidth =9013
                    LayoutCachedHeight =1927
                End
            End
        End
    End
End
