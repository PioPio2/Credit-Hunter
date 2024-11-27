Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4044
    DatasheetFontHeight =10
    ItemSuffix =8
    RecSrcDt = Begin
        0xda23e83f9859e340
    End
    RecordSource ="QueryNewChargebacks"
    Caption ="Sottomaschera QueryNewChargebacks"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="IntestazioneMaschera"
        End
        Begin Section
            Height =1509
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =114
                    Width =900
                    Height =255
                    ColumnWidth =900
                    Name ="Customer_ID"
                    ControlSource ="Customer_ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =114
                            Width =1560
                            Height =255
                            Name ="Customer_ID_Etichetta"
                            Caption ="Customer_ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =456
                    Width =960
                    Height =255
                    ColumnWidth =960
                    TabIndex =1
                    Name ="Document_Number"
                    ControlSource ="Document_Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =456
                            Width =1560
                            Height =255
                            Name ="Document_Number_Etichetta"
                            Caption ="Document_Number"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =798
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="€#,##0.00;-€#,##0.00"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =798
                            Width =1560
                            Height =255
                            Name ="Amount_Etichetta"
                            Caption ="Amount"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1140
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =3
                    Name ="Date"
                    ControlSource ="Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1140
                            Width =1560
                            Height =255
                            Name ="Date_Etichetta"
                            Caption ="Date"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="PièDiPaginaMaschera"
        End
    End
End
