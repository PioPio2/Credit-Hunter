Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
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
    Width =22635
    DatasheetFontHeight =10
    ItemSuffix =180
    Right =28170
    Bottom =13695
    RecSrcDt = Begin
        0x6b3ba7569d3ae640
    End
    RecordSource ="SELECT Tbl_Customers.Name, Tbl_Customers.*, * FROM Tbl_Customers WHERE (((Tbl_Cu"
        "stomers.Credit_controller)=1) AND ((Tbl_Customers.NextAppointment)<=Now())) ORDE"
        "R BY Tbl_Customers.Index DESC; "
    Caption ="Scheduler - Last update: 29/08/2011 08:38:05"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    FilterOnLoad =255
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
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
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
            Height =11448
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Tab
                    OverlapFlags =93
                    BackStyle =0
                    Width =22635
                    Height =11448
                    TabFixedHeight =454
                    Name ="TabCtl48"
                    OnChange ="[Event Procedure]"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedWidth =22635
                    LayoutCachedHeight =11448
                    Begin
                        Begin Page
                            OverlapFlags =87
                            AccessKey =65
                            Left =135
                            Top =585
                            Width =22365
                            Height =10725
                            Name ="PageOverview"
                            Caption ="&All Customers overview"
                            UnicodeAccessKey =65
                            LayoutCachedLeft =135
                            LayoutCachedTop =585
                            LayoutCachedWidth =22500
                            LayoutCachedHeight =11310
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =283
                                    Top =988
                                    Width =21600
                                    Height =9210
                                    Name ="SubmaskSchedulerOverview"
                                    SourceObject ="Form.SubMaskSchedulerOverview"

                                    LayoutCachedLeft =283
                                    LayoutCachedTop =988
                                    LayoutCachedWidth =21883
                                    LayoutCachedHeight =10198
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =285
                                            Top =750
                                            Width =2790
                                            Height =240
                                            Name ="Label176"
                                            Caption ="Scheduler Overview in main currency:"
                                            LayoutCachedLeft =285
                                            LayoutCachedTop =750
                                            LayoutCachedWidth =3075
                                            LayoutCachedHeight =990
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =585
                            Width =22365
                            Height =10725
                            Name ="Sheet1"
                            Caption ="Single &Customer overview"
                            LayoutCachedLeft =135
                            LayoutCachedTop =585
                            LayoutCachedWidth =22500
                            LayoutCachedHeight =11310
                            Begin
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =225
                                    Top =1035
                                    Width =2321
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Testo1"
                                    ControlSource ="Customer_code"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =1035
                                    LayoutCachedWidth =2546
                                    LayoutCachedHeight =1375
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =2805
                                    Top =1035
                                    Width =7460
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Testo3"
                                    ControlSource ="Tbl_Customers.Name"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =2805
                                    LayoutCachedTop =1035
                                    LayoutCachedWidth =10265
                                    LayoutCachedHeight =1375
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =16708
                                    Top =1016
                                    Width =731
                                    Height =340
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Testo5"
                                    ControlSource ="Credit_controller"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =16708
                                    LayoutCachedTop =1016
                                    LayoutCachedWidth =17439
                                    LayoutCachedHeight =1356
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =10433
                                    Top =1014
                                    Width =746
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Testo18"
                                    ControlSource ="Country"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =10433
                                    LayoutCachedTop =1014
                                    LayoutCachedWidth =11179
                                    LayoutCachedHeight =1354
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =84
                                    Left =18138
                                    Top =6519
                                    Width =2486
                                    Height =1301
                                    TabIndex =4
                                    Name ="Comando21"
                                    Caption ="Run s&tatement"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =116
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =18138
                                    LayoutCachedTop =6519
                                    LayoutCachedWidth =20624
                                    LayoutCachedHeight =7820
                                    PictureCaptionArrangement =4
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =9872
                                    Top =8293
                                    Width =7661
                                    Height =2994
                                    TabIndex =5
                                    Name ="Testo10"
                                    HorizontalAnchor =1
                                    VerticalAnchor =2

                                    LayoutCachedLeft =9872
                                    LayoutCachedTop =8293
                                    LayoutCachedWidth =17533
                                    LayoutCachedHeight =11287
                                End
                                Begin Subform
                                    CanGrow = NotDefault
                                    OverlapFlags =247
                                    Left =233
                                    Top =8295
                                    Width =9480
                                    Height =2994
                                    TabIndex =6
                                    Name ="Sottomaschera TblNotes"
                                    SourceObject ="Form.SottomascheraTblNotes"
                                    LinkChildFields ="CustomerCode"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_TblNotes"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2

                                    LayoutCachedLeft =233
                                    LayoutCachedTop =8295
                                    LayoutCachedWidth =9713
                                    LayoutCachedHeight =11289
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    AccessKey =80
                                    IMESentenceMode =3
                                    Left =18138
                                    Top =8934
                                    Width =2486
                                    Height =308
                                    TabIndex =7
                                    Name ="Testo14"
                                    ControlSource ="NextAppointment"
                                    OnLostFocus ="[Event Procedure]"
                                    UnicodeAccessKey =112
                                    HorizontalAnchor =1
                                    ShowDatePicker =0

                                    LayoutCachedLeft =18138
                                    LayoutCachedTop =8934
                                    LayoutCachedWidth =20624
                                    LayoutCachedHeight =9242
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =18138
                                            Top =8537
                                            Width =2486
                                            Height =314
                                            Name ="Etichetta20"
                                            Caption ="Next a&ppointment"
                                            HorizontalAnchor =1
                                            LayoutCachedLeft =18138
                                            LayoutCachedTop =8537
                                            LayoutCachedWidth =20624
                                            LayoutCachedHeight =8851
                                        End
                                    End
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =18138
                                    Top =1979
                                    Width =2486
                                    Height =4067
                                    FontSize =10
                                    TabIndex =8
                                    Name ="Text56"
                                    ControlSource ="DA TOGLIEREEEEEEEE TextEmail"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =18138
                                    LayoutCachedTop =1979
                                    LayoutCachedWidth =20624
                                    LayoutCachedHeight =6046
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =18538
                                    Top =7993
                                    Width =510
                                    Height =355
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =9
                                    Name ="Testo58"
                                    ControlSource ="Index"
                                    Format ="Standard"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =18538
                                    LayoutCachedTop =7993
                                    LayoutCachedWidth =19048
                                    LayoutCachedHeight =8348
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =19908
                                    Top =7993
                                    Width =388
                                    Height =355
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =10
                                    Name ="Testo60"
                                    ControlSource ="Update_date"
                                    Format ="General Date"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =19908
                                    LayoutCachedTop =7993
                                    LayoutCachedWidth =20296
                                    LayoutCachedHeight =8348
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =234
                                    Top =1388
                                    Width =17295
                                    Height =6730
                                    TabIndex =11
                                    Name ="Maschera1"
                                    SourceObject ="Form.MskContainer"
                                    LinkChildFields ="Customer_code"
                                    LinkMasterFields ="Customer_code"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =234
                                    LayoutCachedTop =1388
                                    LayoutCachedWidth =17529
                                    LayoutCachedHeight =8118
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =20907
                                    Top =8923
                                    Width =332
                                    Height =347
                                    TabIndex =12
                                    Name ="Command98"
                                    Caption ="Command98"
                                    OnEnter ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000018000000180000000100180000000000c00600000000000000000000 ,
                                        0x0000000000000000fafbf7fffffcfffefaf8f7f3faf6f1fffffbfff9f5fbf2ee ,
                                        0xede5deddd5cedad4cde6e3dbbab7afe6e4dcddded5e1dfd7fff8f4fff6f2ffff ,
                                        0xfbf8f4effffaf7fffffcfffffcfffffcfdfcf8f4f4eefffffcfffffbfffffaf3 ,
                                        0xede8e5dcd8dacfcbe3d8d4d0c5c1e4dbd7c3bdb8dcd6d1c6c2bddddad5d3cfca ,
                                        0xd3cac6f6ebe7f7eeeafffffbfffffcf2efebfffffcfffffcfffffbf6f4ecffff ,
                                        0xfbfffdf8ece3dfe0d5d1e0d4d0d5c9c5d4c8c4ddd0cecbbebcd7cbc9d5c9c9d3 ,
                                        0xc9c9ccc4c4d4cbc8cdc0bee2d6d2cbc0bcf4ebe7f4efecf2efebfffffcfaf9f5 ,
                                        0xfdfaf2fffdf5f3ede6efe7e0f0e5e1fdf1edfff2efffedecffeeedffedeeffee ,
                                        0xf1f7e6e9fffcffefe0e4fff6faf9ebedfff4f2fff6f3fbefebfdf4f0e6dfdcff ,
                                        0xfffcfffffcf7f6f2fffff8fffff8eae3dafdf4ebfffff9fff1eefff3f2fffcfe ,
                                        0xfff4f8fffafffff8fffff9fffffafffff8fffff9fffffbfffff2f3fff6f3f8ec ,
                                        0xe8fffffbd6cdcaeee9e6faf7f3fffffcf3ede2fdf7ecd6cdc4fff8f0fff8f257 ,
                                        0x4441483332482f33341a20644952472c365238453e24326a515f3218285d4551 ,
                                        0x44323353423f51423ffffffbf4ebe8d5d0cdf2efebfbfaf6fffff5f1e9dcd6cc ,
                                        0xc2f8ebe3fffff945322fdac2c2f5dade92767cdcbfc8e2c4d18c6d7ce8c8d9d6 ,
                                        0xb8cb97788ddcc2d0e5d0d3968582433431fffdf9f2e6e4d0c9c6efeae7fffffc ,
                                        0xf8f2e5cdc5b8e7ddd3f9eae1fffff8291510fffdfcfff4f6a28d90fffbfffff3 ,
                                        0xfca4909cfff9fffff8ff9d8597fff9ffe6cdd1856e6c5c4745fffefbf4e5e2e5 ,
                                        0xdad6cec5c1f7f1ece9e7dce1ddd2d0c6bcf7e8dffffff83c2924c6b8b2c0b7b3 ,
                                        0x777672b8bdbbacb5b26f7776bbbfc0b6b5b7797277c0afb3caadb0ad8d8e6a4c ,
                                        0x4bfff9f8fff9f7e2cec9d7c4bffff1ebe2d8cee8dcd2d1c2b9fffff7fff9f22c ,
                                        0x1f17fffff9f9fbf598a59de7f9f2e7fcf48c9d99effdf9f7fcfba19f9ffffdff ,
                                        0xe9ced1947578513233fff8f7fce3e1ebd5d0ead6d1ddcac3f4d5ccf0d3cae2c8 ,
                                        0xc1ffeee6fffcf438332aeaefe6f5fffa84968feffffef0fffd8e9b99fbfffffb ,
                                        0xf8fa9e9598fff7fbd5c4c8937f845e4b4efffcfefff0efe0cdcaead7d4ead7d2 ,
                                        0xffd3ccf3cec6f0d0caf7e2dafffdf4312f27a9b0a9c6d0ca69726fb1b9b8c2c7 ,
                                        0xc6898b8cbcbbbfbebac0afa8af9b949ba8a1a8726c71494247eee6e7fff8fad9 ,
                                        0xc9cad7c5c4e2cfccf6c8c1eec7bfe6c9c2fffff7fffff8302d25fffffbf1ece9 ,
                                        0xb1a6a8fffcfff8e8efa0939bfffbffd6d8e05a616a3b48500000070410164a52 ,
                                        0x59fbfeffeeeaefe5dadcdfcfd0efdddcffe4dce7c8bfd6bfb7f6e6dffffaf34e ,
                                        0x433ffffdfcfffbfea17c84fff7fffff7ff8e727fc4bac65e6872000910668a92 ,
                                        0xb6d5de99b5bc7e91982e3a40f6faffcec8cdcabdbfe0d0d1e8d6cbd8c8bcf3e7 ,
                                        0xddfcf1e9fffff9190200391113340003540e1b3c00034f0f2229000e1f162300 ,
                                        0x151f5e959a9adadfc1f5fcc7eef7aecdd691a6ae3b454cfafaffe9dee1e1d3d5 ,
                                        0xf6ede3d5ccc2cac5bcfbf3ecfffaf75d3e3ffff7fefff3ffffeefeffeeffffe4 ,
                                        0xf9fff3ffb9b0bd00111a83c7ccacfbfe0010188fbec6cff6ffa3bfc62c3c43fa ,
                                        0xfeffcac0c6f4e6eafaf3eaddd8cfceccc2f1ebe4fffefb2b0c0dca949bd78f9b ,
                                        0xea91a1bf6479d8899ead798b7c738014343a8acfd29eedf0c5ffff11424ad2fb ,
                                        0xffa8c5cc3c4e55f8feffd4cdd2f9eef1fffbf3ebe2d9ddd5cef8f0e9fff3ef49 ,
                                        0x313165393f450913762b397224354d091a5427373c323e0820264f888aa6e9ec ,
                                        0xccffff17454dd4f9ff6e8a912d3d43f8feffcfcaccfffefffffcf8f6e0dbd8c7 ,
                                        0xc4f0e4e0fffdfafff8f6fffcfefff9ffffebf2fff7fffff7fffddee7fffbff78 ,
                                        0x8387446366567f8286aeb388aeb37391963c5257ecfbfee7eceffdfbfbfffcfb ,
                                        0xfffcfbfffcfbecd9d6e1d5d1cac3c0f0ebe8f3e9e9fff6f7fff6f8f2dbdfead3 ,
                                        0xd7fffcfffff7fafafcfdc2cece4c6061516a6e233e42586c71efffffeff8fbd6 ,
                                        0xdbdafdfbfaf9f6f2ffefeffffdfcfff2f1fff2f0dad3d0d4d1cdd0cdc9cac9c5 ,
                                        0xc7c4c0e1dddcd6d1d0d5d0cfbdb8b7e9e5e4fffffefdfffff4fdfff4fffff4fd ,
                                        0xffe1e9e8b6bbbafdfffefaf9f5fffffbfffcfafffefcfff5f3fff5f3f1e8e5d9 ,
                                        0xd2cfe1dcd9d2d1cdd6dad5b0b7b0ccd3ccc6cac4d0d2cccdcac5fdf7f0ede7e2 ,
                                        0xf0eeeefaf9fbe5e5e5dfe0defffffcfcfcf6fffffbfdfef5f6fffefafffeffff ,
                                        0xfefffefefff7f5f6e4e3eadbd8dacfcbc7c3bedcdfd6bfc8beccd7cdcbd5c9cb ,
                                        0xd1c6b0b1a7ebe8e0cec5c2f1e5e5fffaf7fffaf5eeeae5fffff9f1f2e9fdfff6 ,
                                        0xebfffdedfefafbfffefffffefffaf9fffdfcfff1effff8f5f1ebe4cfd2c9d8e2 ,
                                        0xd6bac8bcd1ded0ced8cbe3e7dbdddbd1fffcf8fcedebfffefbfffefafffff9ff ,
                                        0xfff9fafbf1fffff8
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    Picture ="images.bmp"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =20907
                                    LayoutCachedTop =8923
                                    LayoutCachedWidth =21239
                                    LayoutCachedHeight =9270
                                End
                                Begin ComboBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    ColumnCount =4
                                    Left =13313
                                    Top =1017
                                    Width =3221
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =13
                                    ForeColor =255
                                    Name ="Testo99"
                                    ControlSource ="Status"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Tbl_Customer_Status.* FROM Tbl_Customer_Status; "
                                    ColumnWidths ="0;0;0;2268"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =13313
                                    LayoutCachedTop =1017
                                    LayoutCachedWidth =16534
                                    LayoutCachedHeight =1357
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =225
                                    Top =711
                                    Width =1019
                                    Height =231
                                    Name ="Label104"
                                    Caption ="Customer ID"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =711
                                    LayoutCachedWidth =1244
                                    LayoutCachedHeight =942
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =2775
                                    Top =735
                                    Width =1019
                                    Height =231
                                    Name ="Label105"
                                    Caption ="Name"
                                    LayoutCachedLeft =2775
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =3794
                                    LayoutCachedHeight =966
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =10439
                                    Top =690
                                    Width =680
                                    Height =231
                                    Name ="Label106"
                                    Caption ="Country"
                                    HorizontalAnchor =1
                                    LayoutCachedLeft =10439
                                    LayoutCachedTop =690
                                    LayoutCachedWidth =11119
                                    LayoutCachedHeight =921
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =13313
                                    Top =690
                                    Width =680
                                    Height =231
                                    Name ="Label107"
                                    Caption ="Status"
                                    HorizontalAnchor =1
                                    LayoutCachedLeft =13313
                                    LayoutCachedTop =690
                                    LayoutCachedWidth =13993
                                    LayoutCachedHeight =921
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =16673
                                    Top =690
                                    Width =774
                                    Height =231
                                    Name ="Label108"
                                    Caption ="Controller"
                                    HorizontalAnchor =1
                                    LayoutCachedLeft =16673
                                    LayoutCachedTop =690
                                    LayoutCachedWidth =17447
                                    LayoutCachedHeight =921
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =14068
                                    Top =705
                                    Width =2377
                                    Height =219
                                    TabIndex =14
                                    Name ="Text109"
                                    ControlSource ="StatusDate"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =14068
                                    LayoutCachedTop =705
                                    LayoutCachedWidth =16445
                                    LayoutCachedHeight =924
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =18141
                                    Top =1644
                                    Width =2486
                                    Height =231
                                    Name ="Label115"
                                    Caption ="Static notes"
                                    HorizontalAnchor =1
                                    LayoutCachedLeft =18141
                                    LayoutCachedTop =1644
                                    LayoutCachedWidth =20627
                                    LayoutCachedHeight =1875
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =83
                                    Left =18138
                                    Top =9411
                                    Width =2486
                                    Height =1301
                                    TabIndex =15
                                    Name ="Command125"
                                    Caption ="&Save all"
                                    OnClick ="[Event Procedure]"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    UnicodeAccessKey =83
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =18138
                                    LayoutCachedTop =9411
                                    LayoutCachedWidth =20624
                                    LayoutCachedHeight =10712
                                    PictureCaptionArrangement =4
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =11437
                                    Top =1014
                                    Width =1646
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =16
                                    Name ="Text166"
                                    ControlSource ="RetailOEM"
                                    HorizontalAnchor =1

                                    LayoutCachedLeft =11437
                                    LayoutCachedTop =1014
                                    LayoutCachedWidth =13083
                                    LayoutCachedHeight =1354
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =11443
                                    Top =690
                                    Width =680
                                    Height =231
                                    Name ="Label167"
                                    Caption ="Channel"
                                    HorizontalAnchor =1
                                    LayoutCachedLeft =11443
                                    LayoutCachedTop =690
                                    LayoutCachedWidth =12123
                                    LayoutCachedHeight =921
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =72
                            Left =135
                            Top =405
                            Width =22365
                            Height =10905
                            Name ="Sheet2"
                            Caption ="Searc&h"
                            UnicodeAccessKey =104
                            LayoutCachedLeft =135
                            LayoutCachedTop =405
                            LayoutCachedWidth =22500
                            LayoutCachedHeight =11310
                            Begin
                                Begin CommandButton
                                    OverlapFlags =247
                                    AccessKey =83
                                    Left =8250
                                    Top =1515
                                    Height =851
                                    Name ="Comando41"
                                    Caption ="&Search !"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =83

                                    LayoutCachedLeft =8250
                                    LayoutCachedTop =1515
                                    LayoutCachedWidth =9951
                                    LayoutCachedHeight =2366
                                    PictureCaptionArrangement =4
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =285
                                    Top =2010
                                    Width =7773
                                    Height =340
                                    TabIndex =1
                                    Name ="Testo42"

                                    LayoutCachedLeft =285
                                    LayoutCachedTop =2010
                                    LayoutCachedWidth =8058
                                    LayoutCachedHeight =2350
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =285
                                    Top =1665
                                    Width =6480
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Etichetta52"
                                    Caption ="Insert customer name (or part of it) or customer ID or invoice n#"
                                    LayoutCachedLeft =285
                                    LayoutCachedTop =1665
                                    LayoutCachedWidth =6765
                                    LayoutCachedHeight =1905
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =285
                                    Top =735
                                    Width =926
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Testo131"
                                    ControlSource ="Customer_code"

                                    LayoutCachedLeft =285
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =1211
                                    LayoutCachedHeight =1075
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =735
                                    Width =7416
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Testo132"
                                    ControlSource ="Tbl_Customers.Name"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =8856
                                    LayoutCachedHeight =1075
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9165
                                    Top =729
                                    Width =746
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =4
                                    Name ="Testo133"
                                    ControlSource ="Country"

                                    LayoutCachedLeft =9165
                                    LayoutCachedTop =729
                                    LayoutCachedWidth =9911
                                    LayoutCachedHeight =1069
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =285
                                    Top =411
                                    Width =1019
                                    Height =231
                                    Name ="Etichetta134"
                                    Caption ="Customer ID"
                                    LayoutCachedLeft =285
                                    LayoutCachedTop =411
                                    LayoutCachedWidth =1304
                                    LayoutCachedHeight =642
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =1428
                                    Top =411
                                    Width =1019
                                    Height =231
                                    Name ="Etichetta135"
                                    Caption ="Name"
                                    LayoutCachedLeft =1428
                                    LayoutCachedTop =411
                                    LayoutCachedWidth =2447
                                    LayoutCachedHeight =642
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =9171
                                    Top =405
                                    Width =680
                                    Height =231
                                    Name ="Etichetta136"
                                    Caption ="Country"
                                    LayoutCachedLeft =9171
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =9851
                                    LayoutCachedHeight =636
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =405
                            Width =22372
                            Height =10905
                            Name ="Sheet3"
                            Caption ="&E-mail"
                            LayoutCachedLeft =135
                            LayoutCachedTop =405
                            LayoutCachedWidth =22507
                            LayoutCachedHeight =11310
                            Begin
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =225
                                    Top =1650
                                    Width =22272
                                    Height =287
                                    Name ="Testo47"
                                    ControlSource ="Email"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =1650
                                    LayoutCachedWidth =22497
                                    LayoutCachedHeight =1937
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =225
                                    Top =2397
                                    Width =7267
                                    Height =1127
                                    TabIndex =1
                                    Name ="Text45"
                                    ControlSource ="ccEmail"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =2397
                                    LayoutCachedWidth =7492
                                    LayoutCachedHeight =3524
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =225
                                    Top =1333
                                    Width =1298
                                    Height =200
                                    BackColor =-2147483633
                                    Name ="Etichetta53"
                                    Caption ="Main receiver(s)"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =1333
                                    LayoutCachedWidth =1523
                                    LayoutCachedHeight =1533
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =225
                                    Top =2080
                                    Width =1902
                                    Height =258
                                    BackColor =-2147483633
                                    Name ="Etichetta54"
                                    Caption ="External receiver(s) in cc"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =2080
                                    LayoutCachedWidth =2127
                                    LayoutCachedHeight =2338
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =225
                                    Top =735
                                    Width =926
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Testo111"
                                    ControlSource ="Customer_code"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =1151
                                    LayoutCachedHeight =1075
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1360
                                    Top =735
                                    Width =16506
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Testo112"
                                    ControlSource ="Tbl_Customers.Name"

                                    LayoutCachedLeft =1360
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =17866
                                    LayoutCachedHeight =1075
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =225
                                    Top =411
                                    Width =1019
                                    Height =231
                                    Name ="Etichetta113"
                                    Caption ="Customer ID"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =411
                                    LayoutCachedWidth =1244
                                    LayoutCachedHeight =642
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =1360
                                    Top =411
                                    Width =1019
                                    Height =231
                                    Name ="Etichetta114"
                                    Caption ="Name"
                                    LayoutCachedLeft =1360
                                    LayoutCachedTop =411
                                    LayoutCachedWidth =2379
                                    LayoutCachedHeight =642
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =17518
                                    Top =3628
                                    Width =1530
                                    Height =799
                                    TabIndex =4
                                    Name ="Tbl_Templates"
                                    SourceObject ="Form.MskTemplate"
                                    LinkChildFields ="Language"
                                    LinkMasterFields ="Language"

                                    LayoutCachedLeft =17518
                                    LayoutCachedTop =3628
                                    LayoutCachedWidth =19048
                                    LayoutCachedHeight =4427
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =225
                                    Top =5210
                                    Width =22231
                                    Height =4253
                                    TabIndex =5
                                    Name ="Testo155"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =5210
                                    LayoutCachedWidth =22456
                                    LayoutCachedHeight =9463
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =225
                                    Top =4504
                                    Width =22234
                                    Height =305
                                    TabIndex =6
                                    Name ="Testo156"

                                    LayoutCachedLeft =225
                                    LayoutCachedTop =4504
                                    LayoutCachedWidth =22459
                                    LayoutCachedHeight =4809
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =225
                                    Top =4170
                                    Width =1209
                                    Height =231
                                    Name ="Etichetta157"
                                    Caption ="Email subject"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =4170
                                    LayoutCachedWidth =1434
                                    LayoutCachedHeight =4401
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =225
                                    Top =4904
                                    Width =1209
                                    Height =231
                                    Name ="Etichetta158"
                                    Caption ="Email body"
                                    LayoutCachedLeft =225
                                    LayoutCachedTop =4904
                                    LayoutCachedWidth =1434
                                    LayoutCachedHeight =5135
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =8
                                    Left =17985
                                    Top =729
                                    Width =4521
                                    Height =363
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =7
                                    ForeColor =255
                                    Name ="Combo154"
                                    ControlSource ="Status"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT Tbl_Customer_Status.ID, Tbl_Customer_Status.Description, Tbl_Customer_Sta"
                                        "tus.Step, Tbl_Customer_Status.Status, Tbl_Customer_Status.AppearsInTheScheduler,"
                                        " Tbl_Customer_Status.ToSendStatement, Tbl_Customer_Status.ToSendEmail FROM Tbl_C"
                                        "ustomer_Status WHERE (((Tbl_Customer_Status.AppearsInTheScheduler)=Yes)); "
                                    ColumnWidths ="0;2268;0;0;0;0;0;0;0"
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =17985
                                    LayoutCachedTop =729
                                    LayoutCachedWidth =22506
                                    LayoutCachedHeight =1092
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    AccessKey =71
                                    Left =20692
                                    Top =9637
                                    Height =851
                                    TabIndex =8
                                    Name ="Command159"
                                    Caption ="&Go ahead"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =71

                                    LayoutCachedLeft =20692
                                    LayoutCachedTop =9637
                                    LayoutCachedWidth =22393
                                    LayoutCachedHeight =10488
                                    PictureCaptionArrangement =4
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =17985
                                    Top =405
                                    Width =1019
                                    Height =231
                                    Name ="Label160"
                                    Caption ="Status"
                                    LayoutCachedLeft =17985
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =19004
                                    LayoutCachedHeight =636
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =7588
                                    Top =2391
                                    Width =14919
                                    Height =1126
                                    TabIndex =9
                                    Name ="Text162"

                                    LayoutCachedLeft =7588
                                    LayoutCachedTop =2391
                                    LayoutCachedWidth =22507
                                    LayoutCachedHeight =3517
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =7575
                                    Top =2085
                                    Width =2011
                                    Height =258
                                    BackColor =-2147483633
                                    Name ="Label164"
                                    Caption ="Internal receiver(s) in cc"
                                    LayoutCachedLeft =7575
                                    LayoutCachedTop =2085
                                    LayoutCachedWidth =9586
                                    LayoutCachedHeight =2343
                                End
                                Begin Subform
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =16044
                                    Top =3583
                                    Width =1306
                                    Height =857
                                    TabIndex =10
                                    Name ="Sottomaschera Tbl_Link_Customer_Internal_Email_Address"
                                    SourceObject ="Form.Sottomaschera Tbl_Link_Customer_Internal_Email_Address"
                                    LinkChildFields ="CustomerID"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_Tbl_Link_Customer_Internal_Email_Address"

                                    LayoutCachedLeft =16044
                                    LayoutCachedTop =3583
                                    LayoutCachedWidth =17350
                                    LayoutCachedHeight =4440
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =113
                            Top =405
                            Width =22387
                            Height =10905
                            Name ="Sheet 4"
                            EventProcPrefix ="Sheet_4"
                            Caption ="Customer header"
                            LayoutCachedLeft =113
                            LayoutCachedTop =405
                            LayoutCachedWidth =22500
                            LayoutCachedHeight =11310
                            Begin
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =10042
                                    Top =717
                                    Width =837
                                    Height =367
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Text130"
                                    ControlSource ="DSO"
                                    Format ="Standard"

                                    LayoutCachedLeft =10042
                                    LayoutCachedTop =717
                                    LayoutCachedWidth =10879
                                    LayoutCachedHeight =1084
                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =11159
                                    Top =786
                                    Width =6230
                                    Height =3481
                                    TabIndex =1
                                    Name ="Text138"
                                    ControlSource ="Note"

                                    LayoutCachedLeft =11159
                                    LayoutCachedTop =786
                                    LayoutCachedWidth =17389
                                    LayoutCachedHeight =4267
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =10042
                                    Top =405
                                    Width =841
                                    Height =231
                                    Name ="Label139"
                                    Caption ="DSO"
                                    LayoutCachedLeft =10042
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =10883
                                    LayoutCachedHeight =636
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =11145
                                    Top =420
                                    Width =1060
                                    Height =230
                                    Name ="Label140"
                                    Caption ="Report Notes"
                                    LayoutCachedLeft =11145
                                    LayoutCachedTop =420
                                    LayoutCachedWidth =12205
                                    LayoutCachedHeight =650
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =149
                                    Top =735
                                    Width =926
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Text148"
                                    ControlSource ="Customer_code"

                                    LayoutCachedLeft =149
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =1075
                                    LayoutCachedHeight =1075
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1304
                                    Top =735
                                    Width =7416
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Text149"
                                    ControlSource ="Tbl_Customers.Name"

                                    LayoutCachedLeft =1304
                                    LayoutCachedTop =735
                                    LayoutCachedWidth =8720
                                    LayoutCachedHeight =1075
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =9105
                                    Top =734
                                    Width =746
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =4
                                    Name ="Text150"
                                    ControlSource ="Country"

                                    LayoutCachedLeft =9105
                                    LayoutCachedTop =734
                                    LayoutCachedWidth =9851
                                    LayoutCachedHeight =1074
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =149
                                    Top =409
                                    Width =1019
                                    Height =231
                                    Name ="Label151"
                                    Caption ="Customer ID"
                                    LayoutCachedLeft =149
                                    LayoutCachedTop =409
                                    LayoutCachedWidth =1168
                                    LayoutCachedHeight =640
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =1318
                                    Top =409
                                    Width =1019
                                    Height =231
                                    Name ="Label152"
                                    Caption ="Name"
                                    LayoutCachedLeft =1318
                                    LayoutCachedTop =409
                                    LayoutCachedWidth =2337
                                    LayoutCachedHeight =640
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =9111
                                    Top =405
                                    Width =680
                                    Height =231
                                    Name ="Label153"
                                    Caption ="Country"
                                    LayoutCachedLeft =9111
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =9791
                                    LayoutCachedHeight =636
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    AccessKey =79
                                    Left =113
                                    Top =1258
                                    Width =1806
                                    Height =851
                                    TabIndex =5
                                    Name ="Comando137"
                                    Caption ="&Open Master data file"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =79

                                    LayoutCachedLeft =113
                                    LayoutCachedTop =1258
                                    LayoutCachedWidth =1919
                                    LayoutCachedHeight =2109
                                    PictureCaptionArrangement =4
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =679
                    Width =5719
                    Height =272
                    TabIndex =2
                    Name ="Text161"
                    ControlSource ="ccEmail"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    OverlapFlags =255
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18141
                    Top =1020
                    Width =2486
                    Height =340
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text177"
                    ControlSource ="MainPhoneNumber"
                    HorizontalAnchor =1

                    LayoutCachedLeft =18141
                    LayoutCachedTop =1020
                    LayoutCachedWidth =20627
                    LayoutCachedHeight =1360
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =12140
                    Top =1547
                    Width =2992
                    Height =2373
                    TabIndex =1
                    Name ="SubMaskCurrencies"
                    SourceObject ="Form.SubMaskCurrencies"
                    LinkChildFields ="Customer_ID"
                    LinkMasterFields ="Customer_code"

                    LayoutCachedLeft =12140
                    LayoutCachedTop =1547
                    LayoutCachedWidth =15132
                    LayoutCachedHeight =3920
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18141
                    Top =623
                    Width =2375
                    Height =340
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="Text179"
                    ControlSource ="ContactNames"
                    HorizontalAnchor =1

                    LayoutCachedLeft =18141
                    LayoutCachedTop =623
                    LayoutCachedWidth =20516
                    LayoutCachedHeight =963
                End
            End
        End
    End
End
CodeBehindForm
' See "MskScheduler3.cls"
