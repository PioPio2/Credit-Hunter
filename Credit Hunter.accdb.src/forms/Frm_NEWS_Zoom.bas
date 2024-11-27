Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =1
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5555
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =4905
    Top =2310
    Right =10455
    Bottom =6330
    Filter ="[ID]=166"
    RecSrcDt = Begin
        0x2b7356cdb71ee340
    End
    RecordSource ="Tbl_NEWS"
    Caption ="Zoom"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
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
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
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
        Begin FormHeader
            Visible = NotDefault
            Height =0
            BackColor =15523543
            Name ="IntestazioneMaschera"
        End
        Begin Section
            Height =4035
            BackColor =16511715
            Name ="Corpo"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =921
                    Top =690
                    Width =4521
                    BorderColor =9868950
                    ForeColor =26367
                    Name ="TITOLO"
                    ControlSource ="TITOLO"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =690
                            Width =735
                            Height =240
                            FontWeight =700
                            ForeColor =12615680
                            Name ="Etichetta4"
                            Caption ="TITOLO:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =921
                    Top =930
                    Width =4461
                    Height =2913
                    TabIndex =1
                    BorderColor =12615680
                    ForeColor =9868950
                    Name ="NOTE"
                    ControlSource ="NOTE"

                    Begin
                        Begin Label
                            OverlapFlags =95
                            Left =120
                            Top =930
                            Width =705
                            Height =240
                            FontWeight =700
                            ForeColor =12615680
                            Name ="Etichetta5"
                            Caption ="NOTE:"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =255
                    Left =30
                    Top =30
                    Width =5499
                    Height =3975
                    BorderColor =12615680
                    Name ="Casella6"
                End
                Begin Image
                    Left =5209
                    Top =94
                    Width =270
                    Height =270
                    Name ="CmdClose"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000012000000120000000100180000000000f0030000120b0000120b0000 ,
                        0x00000000000000008a62248a6224de08ed6622c5082395091a99091a99091a99 ,
                        0x091a99091a99091a99091a99091a990923956123c3d60eea8a62248a62240000 ,
                        0x8a62247615c80439b11866e30099fd039fff039dff039dff039dff039dff039d ,
                        0xff039dff039fff0399ff0071dc1332ac9406cd8a62240000de08ed0044a90395 ,
                        0xff00aaff01b6ff01b7ff01b7ff01b7ff01b7ff01b7ff01b7ff01b7ff01b7ff01 ,
                        0xb8ff00b1ff387dff0836aecf17ef00005e25c50274de03adff01b7ff00b1ff00 ,
                        0xaeff00b4ff01b4ff01b4ff01b4ff01b4ff00b4ff00aeff00b0ff01b7ff00aeff ,
                        0x0071e06322c600000829980295fc01b4ff00afff8bdeffbbebff00afff00b0ff ,
                        0x01b1ff01b1ff00b0ff00afffbbebff8bdeff00afff01b1ff029bfc082a980000 ,
                        0x081b99039fff00b5ff025ba0cabec5ffffffbde9ff04a7fe00a7fe00a7fe04a7 ,
                        0xfebde9ffffffffcabec5025ba000b5ff03a4ff081c9a0000081c99039bff02a1 ,
                        0xfe00a2f80f4b87c6bbc1ffffffbfe7ff0b9dfe0b9dfebfe7ffffffffc6bbc10f ,
                        0x4b8700a2f802a1fe039aff081c990000091b990491ff0395fe0395fe009aff0f ,
                        0x4587cbc2c4ffffffa7dafea7dafeffffffcbc2c40f4587009aff0395fe0395fe ,
                        0x0490ff091a99000009199a0487ff048afd0489fd048afe008cff104087abaab7 ,
                        0xffffffffffffabaab7104087008cff048afe0489fd048afd0487ff09199a0000 ,
                        0x091799057bff057dfc057cfc057cfc007cfe0378f7abaab7ffffffffffffabaa ,
                        0xb70378f7007cfe057cfc057cfc057dfc057bff09179900000917990670ff0670 ,
                        0xfc0671fc036ffc026dfcc3e0ffffffffb0afbbb0afbbffffffc3e0ff026dfc03 ,
                        0x6ffc0671fc0671fc0670ff0917990000091799066aff0666fb0665fb0666fbbf ,
                        0xd8feffffffc6c0c10e38930e3893c6c0c1ffffffbfd8fe0665fb0666fb0666fb ,
                        0x066aff09179900000916990760ff085ffb0050fbbfd9ffffffffc5bcc1103386 ,
                        0x025efb025dfb103286c5bec1ffffffc0d9ff004ffb075dfb0760ff0916990000 ,
                        0x08269b0a61fc0a7efc0682fd87a6cfb5afb40e4388047dff0b79fd0a74fd0771 ,
                        0xff123b88c3b8b593a4cf025bfc0966fb0a5dfc08249c00006a28cc0951ea038c ,
                        0xfe73e7ff808ba4385d811abdff22b4fd0ea8fc0e9ffc0d95fc0b8eff03398005 ,
                        0x49a20c84fe0572fd0752ea6c23cc0000de08ed295fc80464ff53b5ffe5fffffc ,
                        0xffffc3ffff44f6ff0fe4fe13d9fe12cbfd11befd0eb9ff0ea7ff0584fe2a54ff ,
                        0x0931b9de08ed00008a6224927be20033b7004fe8297afe5184ff5185ff1581ff ,
                        0x0a7eff0b7bff0a77ff0a74ff0a70ff0b6cfe0854e80b2fb8a87ce78a62240000 ,
                        0x8a62248a6224de08ed6a25ce081799080997080a99080a99080b99080c99080d ,
                        0x99080d99080e97091a996d24cede08ed8a62248a62240000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Esci dal programma."
                    Picture ="closeTipNormal.bmp"

                    TabIndex =3
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =921
                    Top =450
                    TabIndex =2
                    BorderColor =9868950
                    ForeColor =16764057
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =120
                            Top =450
                            Width =300
                            Height =240
                            ForeColor =16764057
                            Name ="Etichetta2"
                            Caption ="ID:"
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    Left =120
                    Top =75
                    Width =1680
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Etichetta7"
                    Caption ="Zoom articolo"
                End
                Begin Line
                    OverlapFlags =119
                    Left =56
                    Top =396
                    Width =5386
                    BorderColor =9868950
                    Name ="Linea8"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15523543
            Name ="PièDiPaginaMaschera"
        End
    End
End
CodeBehindForm
' See "Frm_NEWS_Zoom.cls"
