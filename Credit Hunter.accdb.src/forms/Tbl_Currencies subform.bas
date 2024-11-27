Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =0
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6617
    DatasheetFontHeight =11
    ItemSuffix =4
    Left =630
    Top =3465
    Right =4335
    Bottom =6150
    RecSrcDt = Begin
        0x23ada8a2f1d6e340
    End
    RecordSource ="Tbl_Currencies"
    Caption ="Tbl_Currencies subform"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =161
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =161
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =14919545
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =161
            Width =283
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =14919545
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =161
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =14919545
            Name ="FormHeader"
            AutoHeight =1
        End
        Begin Section
            Height =1221
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1815
                    Top =345
                    Width =3660
                    Height =330
                    ColumnWidth =1425
                    Name ="CurrencyID"
                    ControlSource ="CurrencyID"
                    GroupTable =1

                    LayoutCachedLeft =1815
                    LayoutCachedTop =345
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =345
                            Width =1410
                            Height =330
                            Name ="CurrencyID_Label"
                            Caption ="Currency"
                            GroupTable =1
                            LayoutCachedLeft =345
                            LayoutCachedTop =345
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =675
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1815
                    Top =735
                    Width =3660
                    Height =330
                    ColumnWidth =1710
                    TabIndex =1
                    Name ="ExchangeRate"
                    ControlSource ="ExchangeRate"
                    GroupTable =1

                    LayoutCachedLeft =1815
                    LayoutCachedTop =735
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =1065
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =735
                            Width =1410
                            Height =330
                            Name ="ExchangeRate_Label"
                            Caption ="Exchange rate"
                            GroupTable =1
                            LayoutCachedLeft =345
                            LayoutCachedTop =735
                            LayoutCachedWidth =1755
                            LayoutCachedHeight =1065
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
