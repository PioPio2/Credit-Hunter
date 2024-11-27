Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14739
    DatasheetFontHeight =10
    ItemSuffix =58
    Right =18690
    Bottom =13020
    RecSrcDt = Begin
        0x2c03a6745846e340
    End
    RecordSource ="Tbl_Users"
    Caption ="User setup"
    DatasheetFontName ="Arial"
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
        Begin Section
            Height =10261
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2552
                    Top =453
                    Width =2490
                    TabIndex =1
                    Name ="UserName"
                    ControlSource ="UserName"

                    LayoutCachedLeft =2552
                    LayoutCachedTop =453
                    LayoutCachedWidth =5042
                    LayoutCachedHeight =693
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =226
                            Top =450
                            Width =885
                            Height =240
                            Name ="Etichetta3"
                            Caption ="Login name"
                            LayoutCachedLeft =226
                            LayoutCachedTop =450
                            LayoutCachedWidth =1111
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2552
                    Top =793
                    Width =2490
                    TabIndex =2
                    Name ="Name"
                    ControlSource ="Name"

                    LayoutCachedLeft =2552
                    LayoutCachedTop =793
                    LayoutCachedWidth =5042
                    LayoutCachedHeight =1033
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =226
                            Top =793
                            Width =480
                            Height =240
                            Name ="Etichetta5"
                            Caption ="Name"
                            LayoutCachedLeft =226
                            LayoutCachedTop =793
                            LayoutCachedWidth =706
                            LayoutCachedHeight =1033
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =6409
                    Top =396
                    Width =4751
                    Height =2880
                    TabIndex =3
                    Name ="Figlio6"
                    SourceObject ="Table.Tbl_Countries"
                    LinkChildFields ="Credit_controller"
                    LinkMasterFields ="ID"

                    LayoutCachedLeft =6409
                    LayoutCachedTop =396
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =3276
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =225
                    Top =4136
                    Width =4819
                    Height =2955
                    TabIndex =4
                    Name ="Testo7"
                    ControlSource ="EmailText"

                    LayoutCachedLeft =225
                    LayoutCachedTop =4136
                    LayoutCachedWidth =5044
                    LayoutCachedHeight =7091
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5499
                            Top =453
                            Width =765
                            Height =240
                            Name ="Etichetta8"
                            Caption ="Countries"
                            LayoutCachedLeft =5499
                            LayoutCachedTop =453
                            LayoutCachedWidth =6264
                            LayoutCachedHeight =693
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =226
                    Top =3810
                    Width =2835
                    Height =240
                    Name ="Etichetta9"
                    Caption ="Default text email"
                    LayoutCachedLeft =226
                    LayoutCachedTop =3810
                    LayoutCachedWidth =3061
                    LayoutCachedHeight =4050
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =7849
                    TabIndex =5
                    Name ="Check10"
                    ControlSource ="Querywithoutcreditcontroller"

                    LayoutCachedLeft =285
                    LayoutCachedTop =7849
                    LayoutCachedWidth =545
                    LayoutCachedHeight =8089
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =566
                            Top =7849
                            Width =2377
                            Height =231
                            Name ="Label11"
                            Caption ="Query without credit controller"
                            LayoutCachedLeft =566
                            LayoutCachedTop =7849
                            LayoutCachedWidth =2943
                            LayoutCachedHeight =8080
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =8152
                    TabIndex =6
                    Name ="Check12"
                    ControlSource ="Onaccountsstillopen"

                    LayoutCachedLeft =285
                    LayoutCachedTop =8152
                    LayoutCachedWidth =545
                    LayoutCachedHeight =8392
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =568
                            Top =8152
                            Width =1712
                            Height =231
                            Name ="Label13"
                            Caption ="On accounts still open"
                            LayoutCachedLeft =568
                            LayoutCachedTop =8152
                            LayoutCachedWidth =2280
                            LayoutCachedHeight =8383
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =8470
                    TabIndex =7
                    Name ="Check14"
                    ControlSource ="Whopaidyesterdayroutine"

                    LayoutCachedLeft =285
                    LayoutCachedTop =8470
                    LayoutCachedWidth =545
                    LayoutCachedHeight =8710
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =568
                            Top =8470
                            Width =2445
                            Height =231
                            Name ="Label15"
                            Caption ="\"Who paid yesterday\" routine"
                            LayoutCachedLeft =568
                            LayoutCachedTop =8470
                            LayoutCachedWidth =3013
                            LayoutCachedHeight =8701
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3749
                    Top =7845
                    Width =682
                    Height =227
                    TabIndex =10
                    Name ="Text16"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =3749
                    LayoutCachedTop =7845
                    LayoutCachedWidth =4431
                    LayoutCachedHeight =8072
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3751
                    Top =8148
                    Width =682
                    Height =227
                    TabIndex =8
                    Name ="Text18"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =3751
                    LayoutCachedTop =8148
                    LayoutCachedWidth =4433
                    LayoutCachedHeight =8375
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3751
                    Top =8466
                    Width =682
                    Height =227
                    TabIndex =9
                    Name ="Text19"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =3751
                    LayoutCachedTop =8466
                    LayoutCachedWidth =4433
                    LayoutCachedHeight =8693
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4777
                    Top =7845
                    Width =1983
                    Height =227
                    Name ="Combo24"
                    RowSourceType ="Value List"
                    RowSource ="\"Everytime I enter\";\"s\";\"Days\";\"d\";\"Weeks\";\"ww\";\"Months\";\"m\""
                    ColumnWidths ="2268;0"
                    DefaultValue ="\"Everytime I enter\""

                    LayoutCachedLeft =4777
                    LayoutCachedTop =7845
                    LayoutCachedWidth =6760
                    LayoutCachedHeight =8072
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4777
                    Top =8148
                    Width =1983
                    Height =227
                    TabIndex =11
                    Name ="Combo26"
                    RowSourceType ="Value List"
                    RowSource ="\"Everytime I enter\";\"s\";\"Days\";\"d\";\"Weeks\";\"ww\";\"Months\";\"m\""
                    ColumnWidths ="2268;0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"Everytime I enter\""

                    LayoutCachedLeft =4777
                    LayoutCachedTop =8148
                    LayoutCachedWidth =6760
                    LayoutCachedHeight =8375
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4777
                    Top =8466
                    Width =1983
                    Height =227
                    TabIndex =12
                    Name ="Combo27"
                    RowSourceType ="Value List"
                    RowSource ="\"Everytime I enter\";\"s\";\"Days\";\"d\";\"Weeks\";\"ww\";\"Months\";\"m\""
                    ColumnWidths ="2268;0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"Everytime I enter\""

                    LayoutCachedLeft =4777
                    LayoutCachedTop =8466
                    LayoutCachedWidth =6760
                    LayoutCachedHeight =8693
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Top =7845
                    Width =3249
                    Height =227
                    TabIndex =13
                    Name ="Text32"
                    ControlSource ="QuerywithoutcreditcontrollerEvery"
                    DefaultValue ="1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =7845
                    LayoutCachedWidth =10329
                    LayoutCachedHeight =8072
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Top =8148
                    Width =3249
                    Height =227
                    TabIndex =14
                    Name ="Text33"
                    ControlSource ="OnaccountsstillopenEvery"
                    DefaultValue ="1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =8148
                    LayoutCachedWidth =10329
                    LayoutCachedHeight =8375
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Top =8466
                    Width =3249
                    Height =227
                    TabIndex =15
                    Name ="Text34"
                    ControlSource ="WhopaidyesterdayroutineEvery"
                    DefaultValue ="1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =8466
                    LayoutCachedWidth =10329
                    LayoutCachedHeight =8693
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =8830
                    TabIndex =16
                    Name ="Check35"

                    LayoutCachedLeft =285
                    LayoutCachedTop =8830
                    LayoutCachedWidth =545
                    LayoutCachedHeight =9070
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =571
                            Top =8836
                            Width =2689
                            Height =231
                            Name ="Label36"
                            Caption ="Customers with overdue > 30 days"
                            LayoutCachedLeft =571
                            LayoutCachedTop =8836
                            LayoutCachedWidth =3260
                            LayoutCachedHeight =9067
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3751
                    Top =8826
                    Width =682
                    Height =227
                    TabIndex =17
                    Name ="Text37"
                    Format ="Fixed"
                    DefaultValue ="1"

                    LayoutCachedLeft =3751
                    LayoutCachedTop =8826
                    LayoutCachedWidth =4433
                    LayoutCachedHeight =9053
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4777
                    Top =8826
                    Width =1983
                    Height =227
                    TabIndex =18
                    Name ="Combo38"
                    RowSourceType ="Value List"
                    RowSource ="\"Everytime I enter\";\"s\";\"Days\";\"d\";\"Weeks\";\"ww\";\"Months\";\"m\""
                    ColumnWidths ="2268;0"
                    DefaultValue ="\"Everytime I enter\""

                    LayoutCachedLeft =4777
                    LayoutCachedTop =8826
                    LayoutCachedWidth =6760
                    LayoutCachedHeight =9053
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Top =8826
                    Width =3249
                    Height =227
                    TabIndex =19
                    Name ="Text39"
                    DefaultValue ="1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =8826
                    LayoutCachedWidth =10329
                    LayoutCachedHeight =9053
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =285
                    Top =9190
                    TabIndex =20
                    Name ="Check40"

                    LayoutCachedLeft =285
                    LayoutCachedTop =9190
                    LayoutCachedWidth =545
                    LayoutCachedHeight =9430
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =568
                            Top =9190
                            Width =3060
                            Height =245
                            Name ="Label41"
                            Caption ="Customers with overdue > 30 days <0"
                            LayoutCachedLeft =568
                            LayoutCachedTop =9190
                            LayoutCachedWidth =3628
                            LayoutCachedHeight =9435
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3751
                    Top =9186
                    Width =682
                    Height =227
                    TabIndex =21
                    Name ="Text42"
                    Format ="Fixed"
                    DefaultValue ="1"

                    LayoutCachedLeft =3751
                    LayoutCachedTop =9186
                    LayoutCachedWidth =4433
                    LayoutCachedHeight =9413
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4777
                    Top =9186
                    Width =1983
                    Height =227
                    TabIndex =22
                    Name ="Combo43"
                    RowSourceType ="Value List"
                    RowSource ="\"Everytime I enter\";\"s\";\"Days\";\"d\";\"Weeks\";\"ww\";\"Months\";\"m\""
                    ColumnWidths ="2268;0"
                    DefaultValue ="\"Everytime I enter\""

                    LayoutCachedLeft =4777
                    LayoutCachedTop =9186
                    LayoutCachedWidth =6760
                    LayoutCachedHeight =9413
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Top =9186
                    Width =3249
                    Height =227
                    TabIndex =23
                    Name ="Text44"
                    DefaultValue ="1"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =9186
                    LayoutCachedWidth =10329
                    LayoutCachedHeight =9413
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6409
                    Top =4151
                    Width =4804
                    Height =2955
                    TabIndex =24
                    Name ="Testo45"
                    ControlSource ="Signature"

                    LayoutCachedLeft =6409
                    LayoutCachedTop =4151
                    LayoutCachedWidth =11213
                    LayoutCachedHeight =7106
                End
                Begin Label
                    OverlapFlags =93
                    Left =6409
                    Top =3838
                    Width =2835
                    Height =240
                    Name ="Etichetta47"
                    Caption ="Signature"
                    LayoutCachedLeft =6409
                    LayoutCachedTop =3838
                    LayoutCachedWidth =9244
                    LayoutCachedHeight =4078
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2551
                    Top =2379
                    TabIndex =25
                    Name ="Check48"
                    ControlSource ="Superuser"

                    LayoutCachedLeft =2551
                    LayoutCachedTop =2379
                    LayoutCachedWidth =2811
                    LayoutCachedHeight =2619
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =225
                            Top =2379
                            Width =1597
                            Height =231
                            Name ="Label49"
                            Caption ="Superuser (tick=YES)"
                            LayoutCachedLeft =225
                            LayoutCachedTop =2379
                            LayoutCachedWidth =1822
                            LayoutCachedHeight =2610
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =56
                    Top =170
                    Width =11278
                    Height =3312
                    Name ="Box50"
                    LayoutCachedLeft =56
                    LayoutCachedTop =170
                    LayoutCachedWidth =11334
                    LayoutCachedHeight =3482
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2551
                    Top =1605
                    Width =2490
                    TabIndex =26
                    Name ="Text51"
                    ControlSource ="Password"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="Password"

                    LayoutCachedLeft =2551
                    LayoutCachedTop =1605
                    LayoutCachedWidth =5041
                    LayoutCachedHeight =1845
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =225
                            Top =1607
                            Width =1230
                            Height =240
                            Name ="Label52"
                            Caption ="E-mail Password"
                            LayoutCachedLeft =225
                            LayoutCachedTop =1607
                            LayoutCachedWidth =1455
                            LayoutCachedHeight =1847
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2551
                    Top =1965
                    Width =2490
                    TabIndex =27
                    Name ="Text53"
                    ControlSource ="RetypePassword"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="Password"

                    LayoutCachedLeft =2551
                    LayoutCachedTop =1965
                    LayoutCachedWidth =5041
                    LayoutCachedHeight =2205
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =225
                            Top =1967
                            Width =1860
                            Height =240
                            Name ="Label54"
                            Caption ="Re-type E-mail Password"
                            LayoutCachedLeft =225
                            LayoutCachedTop =1967
                            LayoutCachedWidth =2085
                            LayoutCachedHeight =2207
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =60
                    Top =3720
                    Width =11278
                    Height =3702
                    Name ="Box55"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3720
                    LayoutCachedWidth =11338
                    LayoutCachedHeight =7422
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2552
                    Top =1185
                    Width =2490
                    TabIndex =28
                    Name ="Text56"
                    ControlSource ="E-mailAddress"

                    LayoutCachedLeft =2552
                    LayoutCachedTop =1185
                    LayoutCachedWidth =5042
                    LayoutCachedHeight =1425
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =225
                            Top =1187
                            Width =1110
                            Height =240
                            Name ="Label57"
                            Caption ="E-mail address"
                            LayoutCachedLeft =225
                            LayoutCachedTop =1187
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =1427
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskUser.cls"
