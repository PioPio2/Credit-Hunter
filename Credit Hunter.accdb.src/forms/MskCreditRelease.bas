Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =28062
    DatasheetFontHeight =10
    ItemSuffix =130
    Right =28545
    Bottom =13935
    OnUnload ="[Event Procedure]"
    Filter ="(MskCreditRelease.Comment Is Null Or MskCreditRelease.Comment=\"\")"
    RecSrcDt = Begin
        0x719221f1893ae640
    End
    RecordSource ="SELECT DISTINCTROW Tbl_Customers.*, Tbl_PaymentData.PaymentIncoming, Tbl_Payment"
        "Data.PaymentDate, Tbl_PaymentData.Proof, Tbl_PaymentData.Comment, Tbl_CL.Currenc"
        "y, Tbl_CL.CreditLimit, Tbl_CL.OpenARBalance, Tbl_CL.AwaitingInvoicing, Tbl_CL.Am"
        "tScheduledTom, Tbl_CL.AmtScheduled8Dyas, Tbl_Customers.Name, Tbl_Users.*, Tbl_Cu"
        "stomers.Status, Tbl_Customer_Status.Description FROM ((((Tbl_Customers INNER JOI"
        "N Tbl_credit_check_failures ON Tbl_Customers.Customer_code=Tbl_credit_check_fail"
        "ures.[Customer Number]) LEFT JOIN Tbl_PaymentData ON Tbl_Customers.Customer_code"
        "=Tbl_PaymentData.CustomerCode) INNER JOIN Tbl_CL ON Tbl_Customers.Customer_code="
        "Tbl_CL.Customer_code) LEFT JOIN Tbl_Users ON Tbl_Customers.Credit_controller=Tbl"
        "_Users.ID) LEFT JOIN Tbl_Customer_Status ON Tbl_Customers.Status=Tbl_Customer_St"
        "atus.ID ORDER BY Tbl_Customers.Name; "
    Caption ="Release order lines"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnTimer ="[Event Procedure]"
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
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
            Height =13209
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =1530
                    Width =27840
                    Height =11565
                    Name ="TabCtl57"
                    OnChange ="[Event Procedure]"

                    LayoutCachedTop =1530
                    LayoutCachedWidth =27840
                    LayoutCachedHeight =13095
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =1935
                            Width =27570
                            Height =11025
                            Name ="Main form"
                            EventProcPrefix ="Main_form"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1935
                            LayoutCachedWidth =27705
                            LayoutCachedHeight =12960
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =369
                                    Top =2332
                                    Width =4081
                                    Height =283
                                    FontWeight =700
                                    Name ="Testo17"
                                    ControlSource ="Tbl_Customers.Name"

                                    LayoutCachedLeft =369
                                    LayoutCachedTop =2332
                                    LayoutCachedWidth =4450
                                    LayoutCachedHeight =2615
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =369
                                    Top =3099
                                    Width =926
                                    TabIndex =1
                                    Name ="Testo19"
                                    ControlSource ="Tbl_Customers.Customer_code"

                                    LayoutCachedLeft =369
                                    LayoutCachedTop =3099
                                    LayoutCachedWidth =1295
                                    LayoutCachedHeight =3339
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =364
                                            Top =2766
                                            Width =1316
                                            Height =240
                                            Name ="Etichetta49"
                                            Caption ="Customer code:"
                                            LayoutCachedLeft =364
                                            LayoutCachedTop =2766
                                            LayoutCachedWidth =1680
                                            LayoutCachedHeight =3006
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =5830
                                    Top =2326
                                    Height =283
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Testo50"
                                    ControlSource ="RetailOEM"
                                    RowSourceType ="Value List"
                                    RowSource ="Retail;OEM"

                                    LayoutCachedLeft =5830
                                    LayoutCachedTop =2326
                                    LayoutCachedWidth =7531
                                    LayoutCachedHeight =2609
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =5830
                                            Top =2016
                                            Width =1200
                                            Height =240
                                            FontWeight =700
                                            Name ="Etichetta51"
                                            Caption ="Retail/OEM"
                                            LayoutCachedLeft =5830
                                            LayoutCachedTop =2016
                                            LayoutCachedWidth =7030
                                            LayoutCachedHeight =2256
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =5828
                                    Top =3804
                                    Width =1861
                                    Height =283
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Testo41"
                                    RowSourceType ="Value List"
                                    RowSource ="Excellent;Good;Average;Poor"

                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =7689
                                    LayoutCachedHeight =4087
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =5828
                                    Top =3464
                                    Width =1815
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta52"
                                    Caption ="Payment behaviour"
                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =3464
                                    LayoutCachedWidth =7643
                                    LayoutCachedHeight =3704
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2174
                                    Top =3103
                                    Width =991
                                    Height =283
                                    TabIndex =4
                                    Name ="Testo40"
                                    ControlSource ="Country"

                                    LayoutCachedLeft =2174
                                    LayoutCachedTop =3103
                                    LayoutCachedWidth =3165
                                    LayoutCachedHeight =3386
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =5828
                                    Top =3029
                                    Width =2041
                                    Height =283
                                    FontWeight =700
                                    TabIndex =5
                                    Name ="Testo53"
                                    ControlSource ="TotalInsurance"

                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =3029
                                    LayoutCachedWidth =7869
                                    LayoutCachedHeight =3312
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =5828
                                    Top =2757
                                    Width =1624
                                    Height =208
                                    FontWeight =700
                                    Name ="Etichetta54"
                                    Caption ="Insurance CL"
                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =2757
                                    LayoutCachedWidth =7452
                                    LayoutCachedHeight =2965
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =8110
                                    Top =4701
                                    Width =1126
                                    Height =283
                                    FontWeight =700
                                    TabIndex =6
                                    Name ="Testo55"

                                    LayoutCachedLeft =8110
                                    LayoutCachedTop =4701
                                    LayoutCachedWidth =9236
                                    LayoutCachedHeight =4984
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =8110
                                    Top =4230
                                    Width =1577
                                    Height =430
                                    FontWeight =700
                                    Name ="Etichetta56"
                                    Caption ="Exchange rate: 1 EUR = ?? USD"
                                    LayoutCachedLeft =8110
                                    LayoutCachedTop =4230
                                    LayoutCachedWidth =9687
                                    LayoutCachedHeight =4660
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =8083
                                    Top =2027
                                    Width =1891
                                    Height =542
                                    FontWeight =700
                                    Name ="Sottomaschera Tbl_credit_check_failures Etichetta"
                                    Caption ="Payment in transit (in original currency)"
                                    EventProcPrefix ="Sottomaschera_Tbl_credit_check_failures_Etichetta"
                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =2027
                                    LayoutCachedWidth =9974
                                    LayoutCachedHeight =2569
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =8083
                                    Top =2475
                                    Width =1666
                                    Height =283
                                    FontWeight =700
                                    TabIndex =7
                                    Name ="Testo32"
                                    ControlSource ="PaymentIncoming"
                                    Format ="Standard"

                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =2475
                                    LayoutCachedWidth =9749
                                    LayoutCachedHeight =2758
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =8083
                                    Top =3139
                                    Width =1651
                                    Height =272
                                    FontWeight =700
                                    TabIndex =8
                                    Name ="Testo33"
                                    ControlSource ="PaymentDate"

                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =3139
                                    LayoutCachedWidth =9734
                                    LayoutCachedHeight =3411
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =8083
                                    Top =2817
                                    Width =1695
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta36"
                                    Caption ="Payment  date"
                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =2817
                                    LayoutCachedWidth =9778
                                    LayoutCachedHeight =3057
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =8083
                                    Top =3574
                                    Width =1710
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta37"
                                    Caption ="Proof of payment"
                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =3574
                                    LayoutCachedWidth =9793
                                    LayoutCachedHeight =3814
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =10500
                                    Top =2025
                                    Width =1245
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta38"
                                    Caption ="Comment"
                                    LayoutCachedLeft =10500
                                    LayoutCachedTop =2025
                                    LayoutCachedWidth =11745
                                    LayoutCachedHeight =2265
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OldBorderStyle =0
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =8083
                                    Top =3875
                                    Width =1651
                                    Height =285
                                    FontWeight =700
                                    TabIndex =9
                                    Name ="Testo34"
                                    ControlSource ="Proof"
                                    RowSourceType ="Value List"
                                    RowSource ="e-mailed POP;Swift"

                                    LayoutCachedLeft =8083
                                    LayoutCachedTop =3875
                                    LayoutCachedWidth =9734
                                    LayoutCachedHeight =4160
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =10485
                                    Top =2337
                                    Width =9409
                                    Height =2789
                                    FontWeight =700
                                    TabIndex =10
                                    Name ="Testo35"
                                    ControlSource ="Comment"

                                    LayoutCachedLeft =10485
                                    LayoutCachedTop =2337
                                    LayoutCachedWidth =19894
                                    LayoutCachedHeight =5126
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =204
                                    Top =5450
                                    Width =27390
                                    Height =6347
                                    TabIndex =11
                                    Name ="Sottomaschera Tbl_credit_check_failures"
                                    SourceObject ="Form.SottomascheraTbl_credit_check_failures"
                                    LinkChildFields ="Customer Number"
                                    LinkMasterFields ="Tbl_Customers.Customer_code"
                                    EventProcPrefix ="Sottomaschera_Tbl_credit_check_failures"

                                    LayoutCachedLeft =204
                                    LayoutCachedTop =5450
                                    LayoutCachedWidth =27594
                                    LayoutCachedHeight =11797
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =2174
                                    Top =2763
                                    Width =1260
                                    Height =240
                                    Name ="Etichetta61"
                                    Caption ="Country:"
                                    LayoutCachedLeft =2174
                                    LayoutCachedTop =2763
                                    LayoutCachedWidth =3434
                                    LayoutCachedHeight =3003
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =369
                                    Top =2022
                                    Width =1260
                                    Height =240
                                    Name ="Etichetta62"
                                    Caption ="Customer name:"
                                    LayoutCachedLeft =369
                                    LayoutCachedTop =2022
                                    LayoutCachedWidth =1629
                                    LayoutCachedHeight =2262
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =18180
                                    Top =11970
                                    Width =4605
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta86"
                                    Caption ="Total order amount in original currency(USD):"
                                    LayoutCachedLeft =18180
                                    LayoutCachedTop =11970
                                    LayoutCachedWidth =22785
                                    LayoutCachedHeight =12210
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    BorderWidth =3
                                    OverlapFlags =215
                                    TextAlign =3
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =22879
                                    Top =11975
                                    Width =1418
                                    FontWeight =700
                                    TabIndex =12
                                    Name ="Testo85"
                                    ControlSource ="=Format(DSum(\"[amount]\",\"[Tbl_credit_check_failures]\",\"[Customer Number]=\""
                                        " & Testo19.Value),\"##,##0.00\")"

                                    LayoutCachedLeft =22879
                                    LayoutCachedTop =11975
                                    LayoutCachedWidth =24297
                                    LayoutCachedHeight =12215
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =3654
                                    Top =3137
                                    Width =1629
                                    Height =283
                                    TabIndex =13
                                    Name ="Text90"
                                    ControlSource ="Tbl_Users.Name"

                                    LayoutCachedLeft =3654
                                    LayoutCachedTop =3137
                                    LayoutCachedWidth =5283
                                    LayoutCachedHeight =3420
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =3654
                                    Top =2804
                                    Width =1260
                                    Height =240
                                    Name ="Label91"
                                    Caption ="Credit controller"
                                    LayoutCachedLeft =3654
                                    LayoutCachedTop =2804
                                    LayoutCachedWidth =4914
                                    LayoutCachedHeight =3044
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =5828
                                    Top =4625
                                    Width =2065
                                    Height =256
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =14
                                    ForeColor =255
                                    Name ="Testo93"
                                    ControlSource ="Description"

                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =4625
                                    LayoutCachedWidth =7893
                                    LayoutCachedHeight =4881
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =5828
                                    Top =4292
                                    Width =1260
                                    Height =240
                                    Name ="Etichetta94"
                                    Caption ="Status"
                                    LayoutCachedLeft =5828
                                    LayoutCachedTop =4292
                                    LayoutCachedWidth =7088
                                    LayoutCachedHeight =4532
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =22879
                                    Top =12329
                                    Width =1418
                                    Height =240
                                    FontWeight =700
                                    Name ="Etichetta100"
                                    Caption ="223,278.93"
                                    LayoutCachedLeft =22879
                                    LayoutCachedTop =12329
                                    LayoutCachedWidth =24297
                                    LayoutCachedHeight =12569
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =19591
                                    Top =12330
                                    Width =3197
                                    Height =240
                                    FontWeight =700
                                    Name ="Label110"
                                    Caption ="Total order amount in EUR currency: "
                                    LayoutCachedLeft =19591
                                    LayoutCachedTop =12330
                                    LayoutCachedWidth =22788
                                    LayoutCachedHeight =12570
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =370
                                    Top =3804
                                    Width =989
                                    Height =240
                                    Name ="Label113"
                                    Caption ="Oracle CL:"
                                    LayoutCachedLeft =370
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =1359
                                    LayoutCachedHeight =4044
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1426
                                    Top =3804
                                    Width =284
                                    TabIndex =15
                                    Name ="OpenARBalance"
                                    ControlSource ="OpenARBalance"

                                    LayoutCachedLeft =1426
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =1710
                                    LayoutCachedHeight =4044
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3573
                                    Top =3804
                                    TabIndex =16
                                    Name ="CreditLimit"
                                    ControlSource ="CreditLimit"
                                    Format ="Standard"

                                    LayoutCachedLeft =3573
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =5274
                                    LayoutCachedHeight =4044
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1823
                                    Top =3804
                                    Width =284
                                    TabIndex =17
                                    Name ="AmtScheduled8Dyas"
                                    ControlSource ="AmtScheduled8Dyas"

                                    LayoutCachedLeft =1823
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =2107
                                    LayoutCachedHeight =4044
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2220
                                    Top =3804
                                    Width =284
                                    TabIndex =18
                                    Name ="AwaitingInvoicing"
                                    ControlSource ="AwaitingInvoicing"

                                    LayoutCachedLeft =2220
                                    LayoutCachedTop =3804
                                    LayoutCachedWidth =2504
                                    LayoutCachedHeight =4044
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3570
                                    Top =4217
                                    TabIndex =19
                                    Name ="Text114"
                                    ControlSource ="=[OpenARBalance]+[AmtScheduledTom]+[AwaitingInvoicing]"
                                    Format ="Standard"

                                    LayoutCachedLeft =3570
                                    LayoutCachedTop =4217
                                    LayoutCachedWidth =5271
                                    LayoutCachedHeight =4457
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =3570
                                    Top =4651
                                    TabIndex =20
                                    BackColor =255
                                    Name ="Text115"
                                    ControlSource ="=[CreditLimit]-[text114]"
                                    Format ="Standard"

                                    LayoutCachedLeft =3570
                                    LayoutCachedTop =4651
                                    LayoutCachedWidth =5271
                                    LayoutCachedHeight =4891
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =367
                                    Top =4217
                                    Width =3097
                                    Height =245
                                    Name ="Label116"
                                    Caption ="AR bal.+ to be inv. + Scheduled 5 days:"
                                    LayoutCachedLeft =367
                                    LayoutCachedTop =4217
                                    LayoutCachedWidth =3464
                                    LayoutCachedHeight =4462
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =223
                                    Left =367
                                    Top =4651
                                    Width =3097
                                    Height =245
                                    BackColor =255
                                    Name ="Label117"
                                    Caption ="Balance:"
                                    LayoutCachedLeft =367
                                    LayoutCachedTop =4651
                                    LayoutCachedWidth =3464
                                    LayoutCachedHeight =4896
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =283
                                    Top =3590
                                    Width =5216
                                    Height =1532
                                    Name ="Box118"
                                    LayoutCachedLeft =283
                                    LayoutCachedTop =3590
                                    LayoutCachedWidth =5499
                                    LayoutCachedHeight =5122
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =283
                                    Top =1946
                                    Width =5216
                                    Height =1529
                                    Name ="Box119"
                                    LayoutCachedLeft =283
                                    LayoutCachedTop =1946
                                    LayoutCachedWidth =5499
                                    LayoutCachedHeight =3475
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =5719
                                    Top =1962
                                    Width =4431
                                    Height =3159
                                    Name ="Box120"
                                    LayoutCachedLeft =5719
                                    LayoutCachedTop =1962
                                    LayoutCachedWidth =10150
                                    LayoutCachedHeight =5121
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =2607
                                    Top =3760
                                    Width =284
                                    TabIndex =21
                                    Name ="AmtScheduledTom"
                                    ControlSource ="AmtScheduledTom"

                                    LayoutCachedLeft =2607
                                    LayoutCachedTop =3760
                                    LayoutCachedWidth =2891
                                    LayoutCachedHeight =4000
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =23017
                                    Top =2437
                                    Height =851
                                    TabIndex =22
                                    Name ="Command124"
                                    Caption ="Change \"To be released\" status of the current customer"
                                    OnClick ="[Event Procedure]"
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =23017
                                    LayoutCachedTop =2437
                                    LayoutCachedWidth =24718
                                    LayoutCachedHeight =3288
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    OverlapFlags =223
                                    IMESentenceMode =3
                                    Left =20466
                                    Top =3346
                                    Width =6964
                                    Height =1814
                                    FontWeight =700
                                    TabIndex =23
                                    Name ="Text126"
                                    ControlSource ="ReleaseNotes"

                                    LayoutCachedLeft =20466
                                    LayoutCachedTop =3346
                                    LayoutCachedWidth =27430
                                    LayoutCachedHeight =5160
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =20296
                                    Top =2324
                                    Width =7301
                                    Height =2954
                                    Name ="Box127"
                                    LayoutCachedLeft =20296
                                    LayoutCachedTop =2324
                                    LayoutCachedWidth =27597
                                    LayoutCachedHeight =5278
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =1935
                            Width =27570
                            Height =11025
                            Name ="Additional informations"
                            EventProcPrefix ="Additional_informations"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1935
                            LayoutCachedWidth =27705
                            LayoutCachedHeight =12960
                            Begin
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    Left =7935
                                    Top =2455
                                    Width =19625
                                    Height =7132
                                    Name ="Sottomaschera TblNotes"
                                    SourceObject ="Form.SottomascheraTblNotes"
                                    LinkChildFields ="CustomerCode"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_TblNotes"

                                    LayoutCachedLeft =7935
                                    LayoutCachedTop =2455
                                    LayoutCachedWidth =27560
                                    LayoutCachedHeight =9587
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =7935
                                            Top =2115
                                            Width =2631
                                            Height =254
                                            Name ="Etichetta68"
                                            Caption ="Notes"
                                            LayoutCachedLeft =7935
                                            LayoutCachedTop =2115
                                            LayoutCachedWidth =10566
                                            LayoutCachedHeight =2369
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =24888
                                    Top =10091
                                    Width =2566
                                    Height =2677
                                    TabIndex =1
                                    Name ="Comando39"
                                    Caption ="Fill form !"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =24888
                                    LayoutCachedTop =10091
                                    LayoutCachedWidth =27454
                                    LayoutCachedHeight =12768
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =180
                                    Top =10196
                                    Width =14955
                                    Height =2654
                                    TabIndex =2
                                    Name ="MskAging"
                                    SourceObject ="Form.MskAging"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =10196
                                    LayoutCachedWidth =15135
                                    LayoutCachedHeight =12850
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =180
                                            Top =9855
                                            Width =1439
                                            Height =245
                                            Name ="Label95"
                                            Caption ="Aging:"
                                            LayoutCachedLeft =180
                                            LayoutCachedTop =9855
                                            LayoutCachedWidth =1619
                                            LayoutCachedHeight =10100
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1820
                                    Top =2070
                                    Width =3403
                                    Height =283
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Text96"
                                    ControlSource ="Tbl_Customers.Name"

                                    LayoutCachedLeft =1820
                                    LayoutCachedTop =2070
                                    LayoutCachedWidth =5223
                                    LayoutCachedHeight =2353
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =177
                                    Top =2068
                                    Width =1695
                                    Height =240
                                    FontWeight =700
                                    Name ="Label97"
                                    Caption ="Customer name:"
                                    LayoutCachedLeft =177
                                    LayoutCachedTop =2068
                                    LayoutCachedWidth =1872
                                    LayoutCachedHeight =2308
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =178
                                    Top =4809
                                    Width =2241
                                    Height =245
                                    Name ="Etichetta103"
                                    Caption ="CL Excess in EUR currency"
                                    LayoutCachedLeft =178
                                    LayoutCachedTop =4809
                                    LayoutCachedWidth =2419
                                    LayoutCachedHeight =5054
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =3245
                                    Top =4785
                                    Width =1701
                                    Height =240
                                    Name ="Etichetta104"
                                    Caption ="-122,585.39"
                                    LayoutCachedLeft =3245
                                    LayoutCachedTop =4785
                                    LayoutCachedWidth =4946
                                    LayoutCachedHeight =5025
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =3245
                                    Top =5138
                                    Width =1701
                                    Height =240
                                    Name ="Etichetta105"
                                    Caption ="-155,842.09"
                                    LayoutCachedLeft =3245
                                    LayoutCachedTop =5138
                                    LayoutCachedWidth =4946
                                    LayoutCachedHeight =5378
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =340
                                    Top =2566
                                    Width =3844
                                    Height =1523
                                    TabIndex =4
                                    Name ="Tbl_Templates"
                                    SourceObject ="Form.MskTemplate"

                                    LayoutCachedLeft =340
                                    LayoutCachedTop =2566
                                    LayoutCachedWidth =4184
                                    LayoutCachedHeight =4089
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =150
                                    Top =5559
                                    Width =7611
                                    Height =538
                                    FontWeight =700
                                    TabIndex =5
                                    Name ="Testo123"

                                    LayoutCachedLeft =150
                                    LayoutCachedTop =5559
                                    LayoutCachedWidth =7761
                                    LayoutCachedHeight =6097
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =150
                                    Top =6176
                                    Width =7687
                                    Height =3406
                                    FontWeight =700
                                    TabIndex =6
                                    Name ="Testo122"

                                    LayoutCachedLeft =150
                                    LayoutCachedTop =6176
                                    LayoutCachedWidth =7837
                                    LayoutCachedHeight =9582
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =178
                                    Top =5138
                                    Width =2024
                                    Height =244
                                    Name ="Label109"
                                    Caption ="CL Excess in USD currency"
                                    LayoutCachedLeft =178
                                    LayoutCachedTop =5138
                                    LayoutCachedWidth =2202
                                    LayoutCachedHeight =5382
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4233
                    Top =857
                    Width =1941
                    FontWeight =700
                    TabIndex =1
                    ForeColor =255
                    Name ="Testo79"
                    ControlSource ="=DLookUp(\"[Update_Customers_Failing]\",\"[tblgeneral]\")"

                    LayoutCachedLeft =4233
                    LayoutCachedTop =857
                    LayoutCachedWidth =6174
                    LayoutCachedHeight =1097
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =435
                            Top =856
                            Width =3615
                            Height =240
                            Name ="Etichetta80"
                            Caption ="Customers failing the credit check updated as of:"
                            LayoutCachedLeft =435
                            LayoutCachedTop =856
                            LayoutCachedWidth =4050
                            LayoutCachedHeight =1096
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    BorderWidth =3
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4233
                    Top =361
                    Width =1941
                    TabIndex =2
                    Name ="Testo71"
                    ControlSource ="=DLookUp(\"[Update_CL+1]\",\"[tblgeneral]\")"

                    LayoutCachedLeft =4233
                    LayoutCachedTop =361
                    LayoutCachedWidth =6174
                    LayoutCachedHeight =601
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    Left =435
                    Top =360
                    Width =3120
                    Height =240
                    BackColor =-2147483633
                    Name ="Etichetta83"
                    Caption ="Credit limits updated as of:"
                    LayoutCachedLeft =435
                    LayoutCachedTop =360
                    LayoutCachedWidth =3555
                    LayoutCachedHeight =600
                End
                Begin Rectangle
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    Left =177
                    Top =2119
                    Width =26745
                    Height =226
                    Name ="shPB_O2"
                    LayoutCachedLeft =177
                    LayoutCachedTop =2119
                    LayoutCachedWidth =26922
                    LayoutCachedHeight =2345
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =27042
                    Top =2097
                    Width =495
                    Height =270
                    FontWeight =600
                    Name ="Etichetta99"
                    Caption ="100%"
                    LayoutCachedLeft =27042
                    LayoutCachedTop =2097
                    LayoutCachedWidth =27537
                    LayoutCachedHeight =2367
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =86
                    Left =25965
                    Top =340
                    Height =851
                    TabIndex =3
                    Name ="Command125"
                    Caption ="&View report"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =86
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =25965
                    LayoutCachedTop =340
                    LayoutCachedWidth =27666
                    LayoutCachedHeight =1191
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =283
                    Top =120
                    Width =27551
                    Height =1229
                    Name ="Box129"
                    LayoutCachedLeft =283
                    LayoutCachedTop =120
                    LayoutCachedWidth =27834
                    LayoutCachedHeight =1349
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCreditRelease.cls"
