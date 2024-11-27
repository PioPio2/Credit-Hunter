Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11792
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =630
    Top =1665
    Right =18315
    Bottom =10995
    RecSrcDt = Begin
        0xc5c40d428a3ae640
    End
    RecordSource ="SELECT Tbl_Customers.Name, Tbl_Customers.*, Tbl_Customer_Status.ID, Tbl_Customer"
        "_Status.Description, Tbl_Areas.Area, Tbl_Areas.ID, QueryARExposureInMainCurrency"
        ".ARExposure, QueryTotalOverdueOnMonthEndInMainCurrency.TotalOverdue, QueryTotalO"
        "verdueOver90InMainCurrency.TotalOverdueOver90, Tbl_Customers.Index, QueryTotalAl"
        "readyCollectedInEUR.AmountInEUR AS TotalAlreadyCollectedInMainCurrency, * FROM ("
        "((((Tbl_Customer_Status RIGHT JOIN Tbl_Customers ON Tbl_Customer_Status.ID = Tbl"
        "_Customers.Status) LEFT JOIN Tbl_Areas ON Tbl_Customers.Area = Tbl_Areas.ID) LEF"
        "T JOIN QueryARExposureInMainCurrency ON Tbl_Customers.Customer_code = QueryARExp"
        "osureInMainCurrency.Customer_ID) LEFT JOIN QueryTotalOverdueOver90InMainCurrency"
        " ON Tbl_Customers.Customer_code = QueryTotalOverdueOver90InMainCurrency.Customer"
        "_ID) LEFT JOIN QueryTotalOverdueOnMonthEndInMainCurrency ON Tbl_Customers.Custom"
        "er_code = QueryTotalOverdueOnMonthEndInMainCurrency.Customer_ID) LEFT JOIN Query"
        "TotalAlreadyCollectedInEUR ON Tbl_Customers.Customer_code = QueryTotalAlreadyCol"
        "lectedInEUR.CustomerID WHERE (((Tbl_Customers.Credit_controller)=GetNumCreditCon"
        "troller(fOSUserName())) AND ((Tbl_Customers.NextAppointment)<=Now())) ORDER BY T"
        "bl_Customers.Index DESC; "
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
            FontSize =11
            FontName ="Calibri"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =8147
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =453
                    Height =315
                    ColumnWidth =810
                    Name ="ID"
                    ControlSource ="Customer_code"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =453
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =768
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =450
                            Width =1605
                            Height =315
                            Name ="Label1"
                            Caption ="ID"
                            LayoutCachedLeft =113
                            LayoutCachedTop =450
                            LayoutCachedWidth =1718
                            LayoutCachedHeight =765
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =904
                    Height =315
                    ColumnWidth =3525
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Tbl_Customers.Name"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =904
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =1219
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =882
                            Width =1605
                            Height =315
                            Name ="Label3"
                            Caption ="Customer name:"
                            LayoutCachedLeft =113
                            LayoutCachedTop =882
                            LayoutCachedWidth =1718
                            LayoutCachedHeight =1197
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =2274
                    Height =315
                    ColumnWidth =1185
                    TabIndex =2
                    Name ="Channel"
                    ControlSource ="RetailOEM"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =2274
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =2589
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2253
                            Width =1770
                            Height =315
                            Name ="Label9"
                            Caption ="Channel"
                            LayoutCachedLeft =113
                            LayoutCachedTop =2253
                            LayoutCachedWidth =1883
                            LayoutCachedHeight =2568
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =2730
                    Height =315
                    ColumnWidth =3045
                    TabIndex =3
                    Name ="Status"
                    ControlSource ="Description"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =2730
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =3045
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2709
                            Width =1770
                            Height =315
                            Name ="Label11"
                            Caption ="Status"
                            LayoutCachedLeft =113
                            LayoutCachedTop =2709
                            LayoutCachedWidth =1883
                            LayoutCachedHeight =3024
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =3186
                    Height =315
                    ColumnWidth =1665
                    TabIndex =4
                    Name ="Text14"
                    ControlSource ="ARExposure"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =3186
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =3501
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3165
                            Width =1770
                            Height =315
                            Name ="Label15"
                            Caption ="A/R Exposure"
                            LayoutCachedLeft =113
                            LayoutCachedTop =3165
                            LayoutCachedWidth =1883
                            LayoutCachedHeight =3480
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =3642
                    Height =315
                    ColumnWidth =2475
                    TabIndex =5
                    Name ="Text16"
                    ControlSource ="TotalOverdue"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =3642
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =3957
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =3630
                            Width =4290
                            Height =315
                            Name ="Label17"
                            Caption ="Overdue on month end "
                            LayoutCachedLeft =120
                            LayoutCachedTop =3630
                            LayoutCachedWidth =4410
                            LayoutCachedHeight =3945
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =4098
                    Height =315
                    ColumnWidth =2595
                    TabIndex =6
                    Name ="Text18"
                    ControlSource ="TotalOverdueOver90"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =4098
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =4413
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4080
                            Width =3930
                            Height =315
                            Name ="Label19"
                            Caption ="> 90 days on month end"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4080
                            LayoutCachedWidth =4050
                            LayoutCachedHeight =4395
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4536
                    Top =5457
                    Height =315
                    ColumnWidth =2325
                    TabIndex =7
                    Name ="Text28"
                    ControlSource ="=nz([TotalAlreadyCollectedInMainCurrency]/[MonthlyTargetInMainCurrency],0)"
                    Format ="Percent"

                    LayoutCachedLeft =4536
                    LayoutCachedTop =5457
                    LayoutCachedWidth =6237
                    LayoutCachedHeight =5772
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =5439
                            Width =3930
                            Height =315
                            Name ="Label29"
                            Caption ="% Cash target achieved"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5439
                            LayoutCachedWidth =4050
                            LayoutCachedHeight =5754
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SubMaskSchedulerOverview.cls"
