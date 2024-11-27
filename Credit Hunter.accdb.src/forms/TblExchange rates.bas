Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4682
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =1277
    Top =1440
    Right =7241
    Bottom =3206
    RecSrcDt = Begin
        0xd97ed0ceee60e340
    End
    RecordSource ="Tbl_Currencies"
    Caption ="Exchange rates"
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
        Begin Section
            Height =813
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =113
                    Width =965
                    Name ="CurrencyID"
                    ControlSource ="CurrencyID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =883
                            Height =231
                            Name ="Label1"
                            Caption ="Currency"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =453
                    Width =968
                    TabIndex =1
                    Name ="ExchangeRate"
                    ControlSource ="ExchangeRate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =1536
                            Height =231
                            Name ="Label3"
                            Caption ="Exchange Rate"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TblExchange rates.cls"
