Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14910
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =345
    Top =3225
    Right =26850
    Bottom =10020
    Filter ="FiscalYear=2015 AND FiscalQuarter=3 AND FiscalMonth=12"
    RecSrcDt = Begin
        0xd4ba21158a3ae640
    End
    RecordSource ="SELECT Tbl_Cash_Target_Breakdown.CControllerID, Tbl_Cash_Target_Breakdown.Channe"
        "l, Tbl_Cash_Target_Breakdown.OriginalCurrency, Tbl_Cash_Target_Breakdown.FiscalY"
        "ear, Tbl_Cash_Target_Breakdown.FiscalQuarter, Tbl_Cash_Target_Breakdown.FiscalMo"
        "nth, Tbl_Cash_Target_Breakdown.Amount, Tbl_Cash_Target_Breakdown.ChannelCurrency"
        ", Tbl_Cash_Target_Breakdown.ExchangeRateToMainCurrency, Tbl_Cash_Target_Breakdow"
        "n.AmountInUSD FROM Tbl_Cash_Target_Breakdown; "
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
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
        Begin Section
            Height =5952
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =2640
                    Height =315
                    ColumnWidth =1380
                    TabIndex =3
                    Name ="Text8"
                    ControlSource ="FiscalYear"
                    SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                        "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                        "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                        "\015\012\""

                    LayoutCachedLeft =3402
                    LayoutCachedTop =2640
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =2955
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =2640
                            Width =1245
                            Height =315
                            Name ="Label9"
                            Caption ="Fiscal Year"
                            SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                                "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                                "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                                "\015\012\""
                            LayoutCachedLeft =240
                            LayoutCachedTop =2640
                            LayoutCachedWidth =1485
                            LayoutCachedHeight =2955
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =3075
                    Height =315
                    ColumnWidth =1695
                    TabIndex =4
                    Name ="Text10"
                    ControlSource ="FiscalQuarter"
                    SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                        "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                        "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                        "\015\012\""

                    LayoutCachedLeft =3402
                    LayoutCachedTop =3075
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =3390
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =3075
                            Width =1350
                            Height =315
                            Name ="Label11"
                            Caption ="Fiscal Quarter"
                            SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                                "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                                "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                                "\015\012\""
                            LayoutCachedLeft =240
                            LayoutCachedTop =3075
                            LayoutCachedWidth =1590
                            LayoutCachedHeight =3390
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =3510
                    Height =315
                    ColumnWidth =1590
                    TabIndex =5
                    Name ="Text12"
                    ControlSource ="FiscalMonth"

                    LayoutCachedLeft =3402
                    LayoutCachedTop =3510
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =3825
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =3510
                            Width =1245
                            Height =315
                            Name ="Label13"
                            Caption ="Fiscal Month"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3510
                            LayoutCachedWidth =1485
                            LayoutCachedHeight =3825
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =3945
                    Height =315
                    ColumnWidth =3210
                    TabIndex =6
                    Name ="Text14"
                    ControlSource ="Amount"
                    Format ="Standard"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =3402
                    LayoutCachedTop =3945
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =4260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =3945
                            Width =2550
                            Height =315
                            Name ="Label15"
                            Caption ="Amount in billing Currency"
                            LayoutCachedLeft =240
                            LayoutCachedTop =3945
                            LayoutCachedWidth =2790
                            LayoutCachedHeight =4260
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =4380
                    Height =315
                    ColumnWidth =2685
                    TabIndex =7
                    Name ="Text16"
                    ControlSource ="ChannelCurrency"

                    LayoutCachedLeft =3402
                    LayoutCachedTop =4380
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =4695
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =161
                            Left =240
                            Top =4380
                            Width =2340
                            Height =315
                            Name ="Label17"
                            Caption ="Channel Currency Target"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4380
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =4695
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3402
                    Top =1335
                    Height =315
                    ColumnWidth =2505
                    Name ="Text2"
                    ControlSource ="CControllerID"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Users"
                    ColumnWidths ="0;0;1701"
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3402
                    LayoutCachedTop =1335
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =1650
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1335
                            Width =2205
                            Height =315
                            Name ="Label3"
                            Caption ="Credit Controller name"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1335
                            LayoutCachedWidth =2445
                            LayoutCachedHeight =1650
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =2205
                    Height =315
                    ColumnWidth =1875
                    TabIndex =2
                    Name ="Text6"
                    ControlSource ="OriginalCurrency"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Currencies"
                    ColumnWidths ="1701"
                    SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                        "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                        "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                        "\015\012\""
                    AllowValueListEdits =0

                    LayoutCachedLeft =3402
                    LayoutCachedTop =2205
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =2520
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =2198
                            Width =1680
                            Height =315
                            Name ="Label7"
                            Caption ="Billing Currency"
                            SmartTags ="\"Me.MskSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.M"
                                "skSubTarget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = TrueMe.MskSubT"
                                "arget.Form.Filter = Filter\015\012Me.MskSubTarget.visible = True\015\012\015\012"
                                "\015\012\""
                            LayoutCachedLeft =240
                            LayoutCachedTop =2198
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =2513
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3402
                    Top =1770
                    Height =315
                    ColumnWidth =1845
                    TabIndex =1
                    Name ="Text4"
                    ControlSource ="Channel"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Channels"
                    ColumnWidths ="1701;0"
                    OnChange ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3402
                    LayoutCachedTop =1770
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =2085
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1770
                            Width =1365
                            Height =315
                            Name ="Label5"
                            Caption ="Sales Channel"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1770
                            LayoutCachedWidth =1605
                            LayoutCachedHeight =2085
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =4815
                    Height =315
                    ColumnWidth =5025
                    TabIndex =8
                    Name ="Text18"
                    ControlSource ="ExchangeRateToMainCurrency"

                    LayoutCachedLeft =3402
                    LayoutCachedTop =4815
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =5130
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =161
                            Left =240
                            Top =4815
                            Width =2340
                            Height =315
                            Name ="Label19"
                            Caption ="Main Currency/Exchange Rate Billing Currency"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4815
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =5130
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3402
                    Top =5250
                    Height =315
                    ColumnWidth =2265
                    TabIndex =9
                    Name ="Text22"
                    ControlSource ="AmountInUSD"
                    Format ="Standard"

                    LayoutCachedLeft =3402
                    LayoutCachedTop =5250
                    LayoutCachedWidth =5103
                    LayoutCachedHeight =5565
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =161
                            Left =240
                            Top =5250
                            Width =2340
                            Height =315
                            Name ="Label23"
                            Caption ="Target in USD"
                            LayoutCachedLeft =240
                            LayoutCachedTop =5250
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =5565
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskSubTarget2.cls"
