Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4044
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =630
    Top =1350
    Right =17925
    Bottom =5190
    RecSrcDt = Begin
        0x998fe0859759e340
    End
    RecordSource ="QueryOnAccounts"
    Caption ="Sottomaschera QueryOnAccounts"
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
            Height =2834
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =456
                    Width =900
                    Height =255
                    ColumnWidth =1589
                    Name ="Customer_code"
                    ControlSource ="Customer_code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =456
                            Width =1560
                            Height =255
                            Name ="Customer_code_Etichetta"
                            Caption ="Customer ID"
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
                    ColumnWidth =3804
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =798
                            Width =1560
                            Height =255
                            Name ="Name_Etichetta"
                            Caption ="Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1140
                    Width =2310
                    Height =255
                    ColumnWidth =1997
                    TabIndex =2
                    Name ="Customer_reference"
                    ControlSource ="Document_Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1140
                            Width =1560
                            Height =255
                            Name ="Customer_reference_Etichetta"
                            Caption ="Customer reference"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1482
                    Width =1035
                    Height =255
                    ColumnWidth =1141
                    TabIndex =3
                    Name ="Date"
                    ControlSource ="Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1482
                            Width =1560
                            Height =255
                            Name ="Date_Etichetta"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =2187
                    Width =2310
                    Height =255
                    ColumnWidth =1481
                    TabIndex =5
                    Name ="SumOfAmount"
                    ControlSource ="SumOfAmount"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =2187
                            Width =1560
                            Height =255
                            Name ="SumOfAmount_Etichetta"
                            Caption ="Amount"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =1834
                    Width =1035
                    Height =255
                    TabIndex =4
                    Name ="Text12"
                    ControlSource ="Currency"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =1834
                            Width =1560
                            Height =255
                            Name ="Label13"
                            Caption ="Currency"
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
