Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =1
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5693
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =1035
    Top =7005
    Right =5115
    Bottom =8490
    RecSrcDt = Begin
        0x5748ab058a3ae640
    End
    RecordSource ="SELECT Tbl_CashCollected.CustomerID, Tbl_CashCollected.[Payment Date], Tbl_CashC"
        "ollected.Currency, Tbl_CashCollected.Amount, Tbl_CashCollected.[Original amount]"
        " FROM Tbl_CashCollected ORDER BY Tbl_CashCollected.[Payment Date] DESC; "
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin Section
            Height =4478
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1870
                    Top =1700
                    Width =1131
                    Height =315
                    Name ="Text5"
                    ControlSource ="Payment Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =1870
                    LayoutCachedTop =1700
                    LayoutCachedWidth =3001
                    LayoutCachedHeight =2015
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =1695
                            Width =1365
                            Height =315
                            Name ="Label6"
                            Caption ="Payment date"
                            LayoutCachedLeft =165
                            LayoutCachedTop =1695
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2154
                    Top =2154
                    Width =561
                    Height =315
                    ColumnWidth =930
                    TabIndex =1
                    Name ="Text7"
                    ControlSource ="Currency"

                    LayoutCachedLeft =2154
                    LayoutCachedTop =2154
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =2469
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =450
                            Top =2160
                            Width =630
                            Height =315
                            Name ="Label8"
                            Caption ="Currency"
                            LayoutCachedLeft =450
                            LayoutCachedTop =2160
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =2475
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2494
                    Top =2834
                    Width =906
                    Height =315
                    ColumnWidth =1410
                    TabIndex =2
                    Name ="Text9"
                    ControlSource ="Original amount"
                    Format ="Standard"

                    LayoutCachedLeft =2494
                    LayoutCachedTop =2834
                    LayoutCachedWidth =3400
                    LayoutCachedHeight =3149
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =570
                            Top =2835
                            Width =630
                            Height =315
                            Name ="Label10"
                            Caption ="Amount"
                            LayoutCachedLeft =570
                            LayoutCachedTop =2835
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =3150
                        End
                    End
                End
            End
        End
    End
End
