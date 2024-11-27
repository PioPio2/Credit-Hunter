Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14683
    DatasheetFontHeight =10
    ItemSuffix =95
    Right =14505
    Bottom =13095
    RecSrcDt = Begin
        0x49d806408c46e340
    End
    RecordSource ="TblGeneral"
    Caption ="General setup"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Section
            Height =13407
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =623
                    Width =5670
                    Height =255
                    Name ="PathStatemets"
                    ControlSource ="PathStatemets"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =623
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =878
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =623
                            Width =1200
                            Height =240
                            Name ="Label58"
                            Caption ="PathStatemets:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =623
                            LayoutCachedWidth =1484
                            LayoutCachedHeight =863
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =963
                    Width =5670
                    Height =255
                    ColumnWidth =2610
                    TabIndex =1
                    Name ="LastUpdate"
                    ControlSource ="LastUpdate"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =963
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =1218
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =963
                            Width =960
                            Height =240
                            Name ="Label59"
                            Caption ="LastUpdate:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =963
                            LayoutCachedWidth =1244
                            LayoutCachedHeight =1203
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =1303
                    Width =5670
                    Height =255
                    ColumnWidth =5340
                    TabIndex =2
                    Name ="PathLogo"
                    ControlSource ="PathLogo"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =1303
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =1558
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =1303
                            Width =810
                            Height =240
                            Name ="Label60"
                            Caption ="PathLogo:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =1303
                            LayoutCachedWidth =1094
                            LayoutCachedHeight =1543
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =1644
                    Width =5670
                    Height =255
                    ColumnWidth =3097
                    TabIndex =3
                    Name ="PathQueryFile"
                    ControlSource ="PathQueryFile"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =1644
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =1899
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =1644
                            Width =1155
                            Height =240
                            Name ="Label61"
                            Caption ="PathQueryFile:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =1644
                            LayoutCachedWidth =1439
                            LayoutCachedHeight =1884
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =1984
                    Width =5670
                    Height =255
                    ColumnWidth =2934
                    TabIndex =4
                    Name ="PathExcelDirectory"
                    ControlSource ="PathExcelDirectory"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =1984
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =2239
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =1984
                            Width =1500
                            Height =240
                            Name ="Label62"
                            Caption ="PathExcelDirectory:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =1984
                            LayoutCachedWidth =1784
                            LayoutCachedHeight =2224
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =2324
                    Width =5670
                    Height =255
                    TabIndex =5
                    Name ="PathReleases"
                    ControlSource ="PathReleases"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =2324
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =2579
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =2324
                            Width =1110
                            Height =240
                            Name ="Label63"
                            Caption ="PathReleases:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =2324
                            LayoutCachedWidth =1394
                            LayoutCachedHeight =2564
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =2664
                    Width =5670
                    Height =255
                    TabIndex =6
                    Name ="Credit_Manager"
                    ControlSource ="Credit_Manager"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =2664
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =2919
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =2664
                            Width =1290
                            Height =240
                            Name ="Label64"
                            Caption ="Credit_Manager:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =2664
                            LayoutCachedWidth =1574
                            LayoutCachedHeight =2904
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =3004
                    Width =5670
                    Height =255
                    TabIndex =7
                    Name ="Credit_Supervisor"
                    ControlSource ="Credit_Supervisor"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =3004
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =3259
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =3004
                            Width =1425
                            Height =240
                            Name ="Label65"
                            Caption ="Credit_Supervisor:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =3004
                            LayoutCachedWidth =1709
                            LayoutCachedHeight =3244
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =3344
                    Width =5670
                    Height =255
                    ColumnWidth =1956
                    TabIndex =8
                    Name ="Update_CL+1"
                    ControlSource ="Update_CL+1"
                    Format ="General Date"
                    EventProcPrefix ="Update_CL_1"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =3344
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =3599
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =3344
                            Width =1140
                            Height =240
                            Name ="Label66"
                            Caption ="Update_CL+1:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =3344
                            LayoutCachedWidth =1424
                            LayoutCachedHeight =3584
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =3685
                    Width =5670
                    Height =255
                    ColumnWidth =1956
                    TabIndex =9
                    Name ="Update_CL+8"
                    ControlSource ="Update_CL+8"
                    Format ="General Date"
                    EventProcPrefix ="Update_CL_8"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =3685
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =3940
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =3685
                            Width =1140
                            Height =240
                            Name ="Label67"
                            Caption ="Update_CL+8:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =3685
                            LayoutCachedWidth =1424
                            LayoutCachedHeight =3925
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =4025
                    Width =5670
                    Height =255
                    ColumnWidth =2568
                    TabIndex =10
                    Name ="Update_Customers_Failing"
                    ControlSource ="Update_Customers_Failing"
                    Format ="General Date"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =4025
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =4280
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =4025
                            Width =2055
                            Height =240
                            Name ="Label68"
                            Caption ="Update_Customers_Failing:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =4025
                            LayoutCachedWidth =2339
                            LayoutCachedHeight =4265
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =4365
                    Width =5670
                    Height =255
                    TabIndex =11
                    Name ="CreditManagerEmail"
                    ControlSource ="CreditManagerEmail"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =4365
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =4365
                            Width =1560
                            Height =240
                            Name ="Label69"
                            Caption ="CreditManagerEmail:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =4365
                            LayoutCachedWidth =1844
                            LayoutCachedHeight =4605
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =4705
                    Width =5670
                    Height =255
                    TabIndex =12
                    Name ="SupervisorEmail"
                    ControlSource ="SupervisorEmail"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =4705
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =4960
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =4705
                            Width =1260
                            Height =240
                            Name ="Label70"
                            Caption ="SupervisorEmail:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =4705
                            LayoutCachedWidth =1544
                            LayoutCachedHeight =4945
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =5045
                    Width =5670
                    Height =255
                    TabIndex =13
                    Name ="GNT"
                    ControlSource ="GNT"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =5045
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =5300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =5045
                            Width =435
                            Height =240
                            Name ="Label71"
                            Caption ="GNT:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =5045
                            LayoutCachedWidth =719
                            LayoutCachedHeight =5285
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =340
                    Top =5385
                    ColumnWidth =2040
                    TabIndex =14
                    Name ="ImportingProcess"
                    ControlSource ="ImportingProcess"

                    LayoutCachedLeft =340
                    LayoutCachedTop =5385
                    LayoutCachedWidth =600
                    LayoutCachedHeight =5625
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3969
                            Top =5355
                            Width =5670
                            Height =240
                            Name ="Label72"
                            Caption ="ImportingProcess"
                            LayoutCachedLeft =3969
                            LayoutCachedTop =5355
                            LayoutCachedWidth =9639
                            LayoutCachedHeight =5595
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =5725
                    Width =5670
                    Height =255
                    ColumnWidth =2880
                    TabIndex =15
                    Name ="PathChargebackFile"
                    ControlSource ="PathChargebackFile"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =5725
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =5980
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =5725
                            Width =1560
                            Height =240
                            Name ="Label73"
                            Caption ="PathChargebackFile:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =5725
                            LayoutCachedWidth =1844
                            LayoutCachedHeight =5965
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =6066
                    Width =5670
                    Height =255
                    ColumnWidth =2880
                    TabIndex =16
                    Name ="PathWordDirectory"
                    ControlSource ="PathWordDirectory"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =6066
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =6321
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =6066
                            Width =1515
                            Height =240
                            Name ="Label74"
                            Caption ="PathWordDirectory:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =6066
                            LayoutCachedWidth =1799
                            LayoutCachedHeight =6306
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3969
                    Top =6406
                    Width =5670
                    ColumnWidth =3192
                    TabIndex =17
                    Name ="DefaultTemplate"
                    ControlSource ="DefaultTemplate"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Templates.Step, Tbl_Templates.TemplateName FROM Tbl_Templates; "
                    ColumnWidths ="567;3402"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3969
                    LayoutCachedTop =6406
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =7823
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =6406
                            Width =1320
                            Height =240
                            Name ="Label75"
                            Caption ="DefaultTemplate:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =6406
                            LayoutCachedWidth =1604
                            LayoutCachedHeight =6646
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3969
                    Top =7880
                    Width =5670
                    ColumnWidth =3464
                    TabIndex =18
                    Name ="DefaultTemplateInWhoPaidYesterday"
                    ControlSource ="DefaultTemplateInWhoPaidYesterday"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Templates.Step, Tbl_Templates.TemplateName FROM Tbl_Templates; "
                    ColumnWidths ="567;3402"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3969
                    LayoutCachedTop =7880
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =9297
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =7880
                            Width =2835
                            Height =240
                            Name ="Label76"
                            Caption ="DefaultTemplateInWhoPaidYesterday:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =7880
                            LayoutCachedWidth =3119
                            LayoutCachedHeight =8120
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =3969
                    Top =9354
                    Width =5670
                    ColumnWidth =3192
                    TabIndex =19
                    Name ="DefaultTemplateInForwardAging"
                    ControlSource ="DefaultTemplateInForwardAging"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Templates.Step, Tbl_Templates.TemplateName FROM Tbl_Templates; "
                    ColumnWidths ="567;3402"
                    AllowValueListEdits =0

                    LayoutCachedLeft =3969
                    LayoutCachedTop =9354
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =10771
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =9354
                            Width =2475
                            Height =240
                            Name ="Label77"
                            Caption ="DefaultTemplateInForwardAging:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =9354
                            LayoutCachedWidth =2759
                            LayoutCachedHeight =9594
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =10828
                    Width =5670
                    Height =255
                    TabIndex =20
                    Name ="ToBeSentCLto"
                    ControlSource ="ToBeSentCLto"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =10828
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =11083
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =10828
                            Width =1155
                            Height =240
                            Name ="Label78"
                            Caption ="ToBeSentCLto:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =10828
                            LayoutCachedWidth =1439
                            LayoutCachedHeight =11068
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =11168
                    Width =5670
                    Height =255
                    ColumnWidth =2432
                    TabIndex =21
                    Name ="CLHorizonDateLimit"
                    ControlSource ="CLHorizonDateLimit"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =11168
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =11423
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =11168
                            Width =1515
                            Height =240
                            Name ="Label79"
                            Caption ="CLHorizonDateLimit:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =11168
                            LayoutCachedWidth =1799
                            LayoutCachedHeight =11408
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =11508
                    Width =5670
                    Height =255
                    TabIndex =22
                    Name ="Cash Target"
                    ControlSource ="Cash Target"
                    Format ="€#,##0.00;-€#,##0.00"
                    EventProcPrefix ="Cash_Target"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =11508
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =11763
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =11508
                            Width =1020
                            Height =240
                            Name ="Label80"
                            Caption ="Cash Target:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =11508
                            LayoutCachedWidth =1304
                            LayoutCachedHeight =11748
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =11848
                    Width =5670
                    Height =255
                    TabIndex =23
                    Name ="PathInvoiceAttachments"
                    ControlSource ="PathInvoiceAttachments"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =11848
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =12103
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =11848
                            Width =1905
                            Height =240
                            Name ="Label81"
                            Caption ="PathInvoiceAttachments:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =11848
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =12088
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =12188
                    Width =5670
                    Height =816
                    ColumnWidth =5385
                    TabIndex =24
                    Name ="CollectionManagementReportTXTFileHeader"
                    ControlSource ="CollectionManagementReportTXTFileHeader"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =12188
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =13004
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =12188
                            Width =3285
                            Height =240
                            Name ="Label82"
                            Caption ="CollectionManagementReportTXTFileHeader:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =12188
                            LayoutCachedWidth =3569
                            LayoutCachedHeight =12428
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3969
                    Top =13096
                    Width =5670
                    Height =255
                    TabIndex =25
                    Name ="MaxOldStatementInArchive"
                    ControlSource ="MaxOldStatementInArchive"

                    LayoutCachedLeft =3969
                    LayoutCachedTop =13096
                    LayoutCachedWidth =9639
                    LayoutCachedHeight =13351
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =284
                            Top =13096
                            Width =2115
                            Height =240
                            Name ="Label83"
                            Caption ="MaxOldStatementInArchive:"
                            LayoutCachedLeft =284
                            LayoutCachedTop =13096
                            LayoutCachedWidth =2399
                            LayoutCachedHeight =13336
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =900
                    Width =2646
                    Height =255
                    TabIndex =26
                    Name ="Sendusing"
                    ControlSource ="Sendusing"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =900
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =1155
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =900
                            Width =870
                            Height =240
                            Name ="Label84"
                            Caption ="Sendusing:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =900
                            LayoutCachedWidth =11011
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =1240
                    Width =2646
                    Height =255
                    TabIndex =27
                    Name ="SMTPserver"
                    ControlSource ="SMTPserver"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =1240
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =1495
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =1240
                            Width =990
                            Height =240
                            Name ="Label85"
                            Caption ="SMTPserver:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =1240
                            LayoutCachedWidth =11131
                            LayoutCachedHeight =1480
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =1580
                    Width =2646
                    Height =255
                    TabIndex =28
                    Name ="SMTPserverport"
                    ControlSource ="SMTPserverport"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =1580
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =1835
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =1580
                            Width =1290
                            Height =240
                            Name ="Label86"
                            Caption ="SMTPserverport:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =1580
                            LayoutCachedWidth =11431
                            LayoutCachedHeight =1820
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =1920
                    Width =2646
                    Height =255
                    TabIndex =29
                    Name ="SMTPauthenticate"
                    ControlSource ="SMTPauthenticate"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =1920
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =2175
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =1920
                            Width =1440
                            Height =240
                            Name ="Label87"
                            Caption ="SMTPauthenticate:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =1920
                            LayoutCachedWidth =11581
                            LayoutCachedHeight =2160
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11818
                    Top =2260
                    TabIndex =30
                    Name ="SMTPusessl"
                    ControlSource ="SMTPusessl"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =2260
                    LayoutCachedWidth =12078
                    LayoutCachedHeight =2500
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =2211
                            Width =900
                            Height =240
                            Name ="Label88"
                            Caption ="SMTPusessl"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =2211
                            LayoutCachedWidth =11041
                            LayoutCachedHeight =2451
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11818
                    Top =2600
                    Width =2646
                    Height =255
                    TabIndex =31
                    Name ="SMTPconnectiontimeout"
                    ControlSource ="SMTPconnectiontimeout"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =2600
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =2855
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =10141
                            Top =2600
                            Width =1845
                            Height =240
                            Name ="Label89"
                            Caption ="SMTPconnectiontimeout:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =2600
                            LayoutCachedWidth =11986
                            LayoutCachedHeight =2840
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =3798
                    Width =2646
                    Height =255
                    TabIndex =32
                    Name ="MainCurrency"
                    ControlSource ="MainCurrency"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =3798
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =4053
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =3798
                            Width =1125
                            Height =240
                            Name ="Label90"
                            Caption ="MainCurrency:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =3798
                            LayoutCachedWidth =11266
                            LayoutCachedHeight =4038
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =4173
                    Width =2646
                    Height =255
                    TabIndex =33
                    Name ="Text91"
                    ControlSource ="Area"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =4173
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =4428
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10140
                            Top =4170
                            Width =1275
                            Height =240
                            Name ="Label92"
                            Caption ="Logitech Region:"
                            LayoutCachedLeft =10140
                            LayoutCachedTop =4170
                            LayoutCachedWidth =11415
                            LayoutCachedHeight =4410
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11818
                    Top =4548
                    Width =2646
                    Height =1815
                    TabIndex =34
                    Name ="Text93"
                    ControlSource ="ToBeSentCashTargetTo"

                    LayoutCachedLeft =11818
                    LayoutCachedTop =4548
                    LayoutCachedWidth =14464
                    LayoutCachedHeight =6363
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10141
                            Top =4545
                            Width =1305
                            Height =1230
                            Name ="Label94"
                            Caption ="Cash Target Report emails:"
                            LayoutCachedLeft =10141
                            LayoutCachedTop =4545
                            LayoutCachedWidth =11446
                            LayoutCachedHeight =5775
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskTblGeneral.cls"
