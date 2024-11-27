Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7591
    DatasheetFontHeight =11
    ItemSuffix =9
    Right =19200
    Bottom =13095
    RecSrcDt = Begin
        0xf875b956f1d6e340
    End
    Caption ="Cash collected"
    DatasheetFontName ="Calibri"
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
                Begin ComboBox
                    SpecialEffect =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2268
                    Top =859
                    Width =1985
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_MonthEnd.MonthEnd FROM Tbl_MonthEnd ORDER BY Tbl_MonthEnd.MonthEnd; "
                    Format ="Medium Date"

                    LayoutCachedLeft =2268
                    LayoutCachedTop =859
                    LayoutCachedWidth =4253
                    LayoutCachedHeight =1174
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =284
                            Top =855
                            Width =765
                            Height =315
                            Name ="Label3"
                            Caption ="To date"
                            LayoutCachedLeft =284
                            LayoutCachedTop =855
                            LayoutCachedWidth =1049
                            LayoutCachedHeight =1170
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2268
                    Top =345
                    Width =1985
                    Height =315
                    Name ="Text0"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DateAdd(\"d\",[monthend],1) AS FromDate, Tbl_MonthEnd.MonthEnd FROM Tbl_M"
                        "onthEnd ORDER BY DateAdd(\"d\",[monthend],1); "
                    Format ="Medium Date"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2268
                    LayoutCachedTop =345
                    LayoutCachedWidth =4253
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =284
                            Top =349
                            Width =1035
                            Height =315
                            Name ="Label1"
                            Caption ="From date"
                            LayoutCachedLeft =284
                            LayoutCachedTop =349
                            LayoutCachedWidth =1319
                            LayoutCachedHeight =664
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5102
                    Top =3968
                    Height =737
                    TabIndex =2
                    Name ="Command4"
                    Caption ="Run report"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =5102
                    LayoutCachedTop =3968
                    LayoutCachedWidth =6803
                    LayoutCachedHeight =4705
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =255
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2268
                    Top =1704
                    Width =1985
                    Height =315
                    TabIndex =3
                    BoundColumn =1
                    Name ="Combo6"
                    RowSourceType ="Value List"
                    ColumnWidths ="0;1134"
                    AllowValueListEdits =0

                    LayoutCachedLeft =2268
                    LayoutCachedTop =1704
                    LayoutCachedWidth =4253
                    LayoutCachedHeight =2019
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =284
                            Top =1700
                            Width =1635
                            Height =315
                            Name ="Label7"
                            Caption ="Credit Controller"
                            LayoutCachedLeft =284
                            LayoutCachedTop =1700
                            LayoutCachedWidth =1919
                            LayoutCachedHeight =2015
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =283
                    Top =2839
                    Width =3975
                    Height =2955
                    TabIndex =4
                    Name ="Tbl_Currencies subform"
                    SourceObject ="Form.Tbl_Currencies subform"
                    EventProcPrefix ="Tbl_Currencies_subform"

                    LayoutCachedLeft =283
                    LayoutCachedTop =2839
                    LayoutCachedWidth =4258
                    LayoutCachedHeight =5794
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =225
                            Top =2445
                            Width =3780
                            Height =315
                            Name ="Tbl_Currencies subform Label"
                            Caption ="Currencies and Exchange rates used are:"
                            EventProcPrefix ="Tbl_Currencies_subform_Label"
                            LayoutCachedLeft =225
                            LayoutCachedTop =2445
                            LayoutCachedWidth =4005
                            LayoutCachedHeight =2760
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashCollectedByCreditController.cls"
