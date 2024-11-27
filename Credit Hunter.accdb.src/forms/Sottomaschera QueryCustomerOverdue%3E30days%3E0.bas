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
    Width =4100
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =630
    Top =1350
    Right =17925
    Bottom =5190
    RecSrcDt = Begin
        0xb7e6a1319759e340
    End
    RecordSource ="QueryCustomerOverdue>30days>0"
    Caption ="Sottomaschera QueryCustomerOverdue>15days>0"
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
            Height =2789
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Width =900
                    Height =255
                    ColumnWidth =1530
                    Name ="Customer ID"
                    ControlSource ="Customer_code"
                    EventProcPrefix ="Customer_ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Width =1560
                            Height =255
                            Name ="Customer ID_Etichetta"
                            Caption ="Customer ID"
                            EventProcPrefix ="Customer_ID_Etichetta"
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
                    Name ="Customer name"
                    ControlSource ="Customer name"
                    EventProcPrefix ="Customer_name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =798
                            Width =1560
                            Height =255
                            Name ="Customer name_Etichetta"
                            Caption ="Name"
                            EventProcPrefix ="Customer_name_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1674
                    Top =1508
                    Width =2310
                    Height =255
                    ColumnWidth =2985
                    TabIndex =2
                    Name ="Total overdue"
                    ControlSource ="Expr1"
                    Format ="Standard"
                    EventProcPrefix ="Total_overdue"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =54
                            Top =1508
                            Width =1631
                            Height =258
                            Name ="Total overdue_Etichetta"
                            Caption ="Total overdue > 15 days"
                            EventProcPrefix ="Total_overdue_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =2192
                    Width =1035
                    Height =255
                    ColumnWidth =1956
                    TabIndex =4
                    Name ="StatusDate"
                    ControlSource ="StatusDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =2192
                            Width =1560
                            Height =255
                            Name ="StatusDate_Etichetta"
                            Caption ="Status date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =2534
                    Width =900
                    Height =255
                    ColumnWidth =1372
                    TabIndex =5
                    Name ="Days gone by"
                    ControlSource ="Days gone by"
                    EventProcPrefix ="Days_gone_by"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =2534
                            Width =1560
                            Height =255
                            Name ="Days gone by_Etichetta"
                            Caption ="Days gone"
                            EventProcPrefix ="Days_gone_by_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =1850
                    Width =2310
                    Height =255
                    ColumnWidth =2160
                    TabIndex =3
                    Name ="Status"
                    ControlSource ="Status"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =1850
                            Width =1560
                            Height =255
                            Name ="Status_Etichetta"
                            Caption ="Status"
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
