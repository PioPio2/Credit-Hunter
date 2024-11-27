Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5103
    DatasheetFontHeight =10
    ItemSuffix =14
    Right =10164
    Bottom =10080
    RecSrcDt = Begin
        0xd432f017b555e340
    End
    RecordSource ="Tbl_Customer_Status"
    Caption ="Customer statuses"
    DatasheetFontName ="Arial"
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
        Begin Section
            Height =3536
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =453
                    Width =3119
                    ColumnWidth =3423
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =109
                            Top =448
                            Width =1698
                            Height =775
                            Name ="Label3"
                            Caption ="Description (appears in the first and in the third page of the scheduler)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =1494
                    Width =3119
                    TabIndex =1
                    Name ="Text4"
                    ControlSource ="Status"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =109
                            Top =1496
                            Width =1696
                            Height =589
                            Name ="Label5"
                            Caption ="Status (appears in the first page of the scheduler only)"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =190
                    Top =2516
                    TabIndex =2
                    Name ="Check8"
                    ControlSource ="AppearsInTheScheduler"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =421
                            Top =2486
                            Width =3967
                            Height =231
                            Name ="Label9"
                            Caption ="Status appears in the scheduler list"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =190
                    Top =2906
                    TabIndex =3
                    Name ="Check10"
                    ControlSource ="ToSendStatement"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =421
                            Top =2880
                            Width =1766
                            Height =231
                            Name ="Label11"
                            Caption ="Statement will be sent"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =190
                    Top =3296
                    TabIndex =4
                    Name ="Check12"
                    ControlSource ="ToSendEmail"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =421
                            Top =3261
                            Width =1331
                            Height =231
                            Name ="Label13"
                            Caption ="Email will be sent"
                        End
                    End
                End
            End
        End
    End
End
