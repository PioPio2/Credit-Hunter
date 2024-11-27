Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7088
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =24210
    Top =1185
    Right =25575
    Bottom =2190
    Filter ="Step = 0"
    RecSrcDt = Begin
        0x9e5013265d58e340
    End
    RecordSource ="Tbl_Templates"
    Caption ="Tbl_Templates"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
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
            Height =3841
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =851
                    Top =737
                    Width =6237
                    Height =2730
                    Name ="Testo3"
                    ControlSource ="Text"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =737
                            Width =405
                            Height =240
                            Name ="Etichetta5"
                            Caption ="Text"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =851
                    Top =113
                    Width =6237
                    Height =225
                    TabIndex =1
                    Name ="Testo1"
                    ControlSource ="Subject"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =113
                            Width =615
                            Height =240
                            Name ="Etichetta11"
                            Caption ="Subject"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =851
                    Top =405
                    Width =507
                    Height =255
                    TabIndex =2
                    Name ="Testo2"
                    ControlSource ="Step"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =405
                            Width =615
                            Height =240
                            Name ="Etichetta13"
                            Caption ="Step"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2780
                    Top =394
                    Width =507
                    Height =255
                    TabIndex =3
                    Name ="Text14"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1929
                            Top =394
                            Width =615
                            Height =240
                            Name ="Label15"
                            Caption ="ID"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskTemplate.cls"
