Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =27892
    DatasheetFontHeight =10
    ItemSuffix =12
    Right =28800
    Bottom =14190
    RecSrcDt = Begin
        0xff13e7719759e340
    End
    Caption ="Dashboard"
    OnOpen ="[Event Procedure]"
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Tab
            Width =5103
            Height =3402
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =13096
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Width =27630
                    Height =12990
                    Name ="TabCtl5"

                    LayoutCachedWidth =27630
                    LayoutCachedHeight =12990
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Top =396
                            Width =27495
                            Height =12465
                            Name ="Main dashboard"
                            EventProcPrefix ="Main_dashboard"
                            LayoutCachedTop =396
                            LayoutCachedWidth =27495
                            LayoutCachedHeight =12861
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Top =396
                                    Width =27495
                                    Height =12465
                                    Name ="Dashboard"
                                    SourceObject ="Form.MskDashboard"

                                    LayoutCachedTop =396
                                    LayoutCachedWidth =27495
                                    LayoutCachedHeight =12861
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =405
                            Width =27360
                            Height =12450
                            Name ="On account amounts"
                            EventProcPrefix ="On_account_amounts"
                            LayoutCachedLeft =135
                            LayoutCachedTop =405
                            LayoutCachedWidth =27495
                            LayoutCachedHeight =12855
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =283
                                    Top =720
                                    Width =17577
                                    Height =4125
                                    Name ="Sottomaschera QueryOnAccounts"
                                    SourceObject ="Form.Sottomaschera QueryOnAccounts"
                                    EventProcPrefix ="Sottomaschera_QueryOnAccounts"

                                    LayoutCachedLeft =283
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =17860
                                    LayoutCachedHeight =4845
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =405
                            Width =27360
                            Height =12450
                            Name ="Customers with overdue >15 days"
                            EventProcPrefix ="Customers_with_overdue__15_days"
                            LayoutCachedLeft =135
                            LayoutCachedTop =405
                            LayoutCachedWidth =27495
                            LayoutCachedHeight =12855
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =283
                                    Top =720
                                    Width =17577
                                    Height =4125
                                    Name ="Figlio3"
                                    SourceObject ="Form.Sottomaschera QueryCustomerOverdue>30days>0"

                                    LayoutCachedLeft =283
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =17860
                                    LayoutCachedHeight =4845
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =405
                            Width =27360
                            Height =12450
                            Name ="Customers without overdue > 30 days"
                            EventProcPrefix ="Customers_without_overdue___30_days"
                            LayoutCachedLeft =135
                            LayoutCachedTop =405
                            LayoutCachedWidth =27495
                            LayoutCachedHeight =12855
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =290
                                    Top =720
                                    Width =17577
                                    Height =4125
                                    Name ="Sottomaschera QueryCustomerOverdue>30days<0"
                                    SourceObject ="Form.Sottomaschera QueryCustomerOverdue>30days<0"
                                    EventProcPrefix ="Sottomaschera_QueryCustomerOverdue_30days_0"

                                    LayoutCachedLeft =290
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =17867
                                    LayoutCachedHeight =4845
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Maschera1.cls"
