Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13209
    DatasheetFontHeight =10
    ItemSuffix =28
    Right =10164
    Bottom =10080
    RecSrcDt = Begin
        0x4fc1b9677babe340
    End
    RecordSource ="SELECT Tbl_Banks.Country, Tbl_Banks.* FROM Tbl_Banks ORDER BY Tbl_Banks.Country;"
        " "
    Caption ="Logitech bank details"
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2777
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =1185
                    Top =245
                    Width =2835
                    ColumnWidth =855
                    Name ="Country"
                    ControlSource ="Tbl_Banks.Country"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Countries.Code FROM Tbl_Countries; "
                    ColumnWidths ="1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =245
                            Width =828
                            Height =240
                            Name ="Etichetta3"
                            Caption ="Country"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1181
                    Top =790
                    Width =2835
                    ColumnWidth =2775
                    TabIndex =1
                    Name ="USDLine1"
                    ControlSource ="USDLine1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =790
                            Width =828
                            Height =244
                            Name ="Etichetta5"
                            Caption ="USD Line1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1181
                    Top =1130
                    Width =2835
                    ColumnWidth =2565
                    TabIndex =2
                    Name ="USDLine2"
                    ControlSource ="USDLine2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1128
                            Width =828
                            Height =244
                            Name ="Etichetta7"
                            Caption ="USD Line2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1181
                    Top =1471
                    Width =2835
                    ColumnWidth =3045
                    TabIndex =3
                    Name ="USDLine3"
                    ControlSource ="USDLine3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1467
                            Width =828
                            Height =245
                            Name ="Etichetta9"
                            Caption ="USD Line3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1181
                    Top =1811
                    Width =2835
                    TabIndex =4
                    Name ="USDLine4"
                    ControlSource ="USDLine4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1811
                            Width =828
                            Height =240
                            Name ="Etichetta11"
                            Caption ="USD Line4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5461
                    Top =788
                    Width =2835
                    ColumnWidth =3060
                    TabIndex =5
                    Name ="EURLine1"
                    ControlSource ="EURLine1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4592
                            Top =788
                            Width =750
                            Height =240
                            Name ="Etichetta13"
                            Caption ="EUR Line1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5461
                    Top =1128
                    Width =2835
                    ColumnWidth =3450
                    TabIndex =6
                    Name ="EURLine2"
                    ControlSource ="EURLine2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4592
                            Top =1128
                            Width =750
                            Height =240
                            Name ="Etichetta15"
                            Caption ="EUR Line2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5461
                    Top =1469
                    Width =2835
                    ColumnWidth =2610
                    TabIndex =7
                    Name ="EURLine3"
                    ControlSource ="EURLine3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4592
                            Top =1469
                            Width =750
                            Height =240
                            Name ="Etichetta17"
                            Caption ="EUR Line3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5461
                    Top =1809
                    Width =2835
                    TabIndex =8
                    Name ="EURLine4"
                    ControlSource ="EURLine4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4592
                            Top =1809
                            Width =750
                            Height =240
                            Name ="Etichetta19"
                            Caption ="EUR Line4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9645
                    Top =788
                    Width =2835
                    Height =267
                    TabIndex =9
                    Name ="Text20"
                    ControlSource ="GBPLine1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8777
                            Top =788
                            Width =750
                            Height =240
                            Name ="Label21"
                            Caption ="GBP Line1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9645
                    Top =1128
                    Width =2835
                    Height =267
                    TabIndex =10
                    Name ="Text22"
                    ControlSource ="GBPLine2"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8777
                            Top =1128
                            Width =750
                            Height =240
                            Name ="Label23"
                            Caption ="GBP Line2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9645
                    Top =1469
                    Width =2835
                    Height =267
                    TabIndex =11
                    Name ="Text24"
                    ControlSource ="GBPLine3"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8777
                            Top =1469
                            Width =750
                            Height =240
                            Name ="Label25"
                            Caption ="GBP Line3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9645
                    Top =1809
                    Width =2835
                    Height =267
                    TabIndex =12
                    Name ="Text26"
                    ControlSource ="GBPLine4"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8777
                            Top =1809
                            Width =750
                            Height =240
                            Name ="Label27"
                            Caption ="GBP Line4"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskTblBanks.cls"
