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
    Width =4561
    DatasheetFontHeight =10
    ItemSuffix =4
    Right =15360
    Bottom =10056
    RecSrcDt = Begin
        0xe35e9a7f4058e340
    End
    RecordSource ="Tbl_Languages"
    Caption ="Template language names"
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
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =240
                            Height =240
                            Name ="Etichetta1"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    TabIndex =1
                    Name ="Language"
                    ControlSource ="Language"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =780
                            Height =240
                            Name ="Etichetta3"
                            Caption ="Language"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskLanguages.cls"
