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
    Width =26985
    DatasheetFontHeight =11
    ItemSuffix =56
    Right =28545
    Bottom =12300
    RecSrcDt = Begin
        0x233366dcb3eee340
    End
    Caption ="Cash Targets"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            Height =10658
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =345
                    Top =3221
                    Width =26520
                    Height =7095
                    Name ="MskSubTarget"
                    SourceObject ="Form.MskSubTarget2"

                    LayoutCachedLeft =345
                    LayoutCachedTop =3221
                    LayoutCachedWidth =26865
                    LayoutCachedHeight =10316
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =345
                            Top =2790
                            Width =1755
                            Height =315
                            Name ="Label0"
                            Caption ="Target Breakdown"
                            LayoutCachedLeft =345
                            LayoutCachedTop =2790
                            LayoutCachedWidth =2100
                            LayoutCachedHeight =3105
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =345
                    Top =649
                    Height =315
                    TabIndex =1
                    Name ="Text1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_MonthEnd.FiscalYear FROM Tbl_MonthEnd GROUP BY Tbl_MonthEnd.FiscalYea"
                        "r; "
                    OnChange ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =345
                    LayoutCachedTop =649
                    LayoutCachedWidth =2046
                    LayoutCachedHeight =964
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =345
                            Top =196
                            Width =1035
                            Height =315
                            Name ="Label2"
                            Caption ="Fiscal Year"
                            LayoutCachedLeft =345
                            LayoutCachedTop =196
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =511
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2955
                    Top =648
                    Height =315
                    TabIndex =2
                    Name ="Text3"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4"
                    OnChange ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2955
                    LayoutCachedTop =648
                    LayoutCachedWidth =4656
                    LayoutCachedHeight =963
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =2955
                            Top =195
                            Width =1350
                            Height =315
                            Name ="Label4"
                            Caption ="Fiscal Quarter"
                            LayoutCachedLeft =2955
                            LayoutCachedTop =195
                            LayoutCachedWidth =4305
                            LayoutCachedHeight =510
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5100
                    Top =648
                    Height =315
                    TabIndex =3
                    Name ="Text5"
                    RowSourceType ="Value List"
                    OnChange ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =5100
                    LayoutCachedTop =648
                    LayoutCachedWidth =6801
                    LayoutCachedHeight =963
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =5100
                            Top =195
                            Width =1560
                            Height =315
                            Name ="Label6"
                            Caption ="Calendar Month"
                            LayoutCachedLeft =5100
                            LayoutCachedTop =195
                            LayoutCachedWidth =6660
                            LayoutCachedHeight =510
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9232
                    Top =745
                    Width =1375
                    TabIndex =4
                    Name ="Check42"

                    LayoutCachedLeft =9232
                    LayoutCachedTop =745
                    LayoutCachedWidth =10607
                    LayoutCachedHeight =985
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9547
                            Top =715
                            Width =1785
                            Height =315
                            Name ="Label43"
                            Caption ="Monday"
                            LayoutCachedLeft =9547
                            LayoutCachedTop =715
                            LayoutCachedWidth =11332
                            LayoutCachedHeight =1030
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9232
                    Top =1180
                    Width =1375
                    TabIndex =5
                    Name ="Check44"

                    LayoutCachedLeft =9232
                    LayoutCachedTop =1180
                    LayoutCachedWidth =10607
                    LayoutCachedHeight =1420
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9547
                            Top =1150
                            Width =1785
                            Height =315
                            Name ="Label45"
                            Caption ="Tuesday"
                            LayoutCachedLeft =9547
                            LayoutCachedTop =1150
                            LayoutCachedWidth =11332
                            LayoutCachedHeight =1465
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9232
                    Top =1615
                    Width =1375
                    TabIndex =6
                    Name ="Check46"

                    LayoutCachedLeft =9232
                    LayoutCachedTop =1615
                    LayoutCachedWidth =10607
                    LayoutCachedHeight =1855
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9542
                            Top =1585
                            Width =2115
                            Height =315
                            Name ="Label47"
                            Caption ="Wednesday"
                            LayoutCachedLeft =9542
                            LayoutCachedTop =1585
                            LayoutCachedWidth =11657
                            LayoutCachedHeight =1900
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9232
                    Top =2050
                    Width =1375
                    TabIndex =7
                    Name ="Check48"

                    LayoutCachedLeft =9232
                    LayoutCachedTop =2050
                    LayoutCachedWidth =10607
                    LayoutCachedHeight =2290
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9542
                            Top =2020
                            Width =1860
                            Height =315
                            Name ="Label49"
                            Caption ="Thursday"
                            LayoutCachedLeft =9542
                            LayoutCachedTop =2020
                            LayoutCachedWidth =11402
                            LayoutCachedHeight =2335
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =9232
                    Top =2485
                    Width =1375
                    TabIndex =8
                    Name ="Check50"

                    LayoutCachedLeft =9232
                    LayoutCachedTop =2485
                    LayoutCachedWidth =10607
                    LayoutCachedHeight =2725
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =9547
                            Top =2455
                            Width =1785
                            Height =315
                            Name ="Label51"
                            Caption ="Friday"
                            LayoutCachedLeft =9547
                            LayoutCachedTop =2455
                            LayoutCachedWidth =11332
                            LayoutCachedHeight =2770
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =8955
                    Top =623
                    Width =3540
                    Height =2278
                    Name ="Box54"
                    LayoutCachedLeft =8955
                    LayoutCachedTop =623
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =2901
                End
                Begin Label
                    OverlapFlags =85
                    Left =8957
                    Top =226
                    Width =2535
                    Height =285
                    Name ="Label55"
                    Caption ="Email to be sent on"
                    LayoutCachedLeft =8957
                    LayoutCachedTop =226
                    LayoutCachedWidth =11492
                    LayoutCachedHeight =511
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashTargets2.cls"
