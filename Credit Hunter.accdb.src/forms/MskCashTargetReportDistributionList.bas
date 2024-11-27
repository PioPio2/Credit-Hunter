Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11168
    DatasheetFontHeight =10
    ItemSuffix =10
    Right =14505
    Bottom =13095
    RecSrcDt = Begin
        0x40e6f00e36f3e340
    End
    RecordSource ="TblGeneral"
    Caption ="Email Cash Target Distribution Lists"
    DatasheetFontName ="Arial"
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
            Height =7086
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =405
                    Top =836
                    Width =7938
                    Height =1695
                    Name ="ID"
                    ControlSource ="ToBeSentCashTargetTo"

                    LayoutCachedLeft =405
                    LayoutCachedTop =836
                    LayoutCachedWidth =8343
                    LayoutCachedHeight =2531
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =405
                            Top =390
                            Width =5160
                            Height =270
                            Name ="Etichetta1"
                            Caption ="Email Main Recipients who will receive the report every day"
                            LayoutCachedLeft =405
                            LayoutCachedTop =390
                            LayoutCachedWidth =5565
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =404
                    Top =4952
                    Width =7938
                    Height =1695
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="ToBeSentCashTargetToSecondGroup"

                    LayoutCachedLeft =404
                    LayoutCachedTop =4952
                    LayoutCachedWidth =8342
                    LayoutCachedHeight =6647
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =404
                            Top =4499
                            Width =7710
                            Height =300
                            Name ="Label5"
                            Caption ="Email Main Recipients who will receive the report starting from xxx days before "
                                "the fiscal month end."
                            LayoutCachedLeft =404
                            LayoutCachedTop =4499
                            LayoutCachedWidth =8114
                            LayoutCachedHeight =4799
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2270
                    Top =3870
                    Width =891
                    Height =255
                    TabIndex =1
                    Name ="Text6"
                    ControlSource ="CashCollectedSecondGroupDays"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =2270
                    LayoutCachedTop =3870
                    LayoutCachedWidth =3161
                    LayoutCachedHeight =4125
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =405
                            Top =3871
                            Width =1170
                            Height =240
                            Name ="Label7"
                            Caption ="Threshold days"
                            LayoutCachedLeft =405
                            LayoutCachedTop =3871
                            LayoutCachedWidth =1575
                            LayoutCachedHeight =4111
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =225
                    Top =3585
                    Width =8622
                    Height =3232
                    Name ="Box8"
                    LayoutCachedLeft =225
                    LayoutCachedTop =3585
                    LayoutCachedWidth =8847
                    LayoutCachedHeight =6817
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =225
                    Top =135
                    Width =8622
                    Height =2707
                    Name ="Box9"
                    LayoutCachedLeft =225
                    LayoutCachedTop =135
                    LayoutCachedWidth =8847
                    LayoutCachedHeight =2842
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashTargetReportDistributionList.cls"
