Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2769
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =2490
    Top =10485
    Right =5805
    Bottom =11655
    RecSrcDt = Begin
        0x186c73d29d3ae640
    End
    RecordSource ="SELECT Tbl_Invoices.Currency, Tbl_Invoices.Customer_ID, Tbl_Invoices.Update_date"
        " FROM Tbl_Invoices GROUP BY Tbl_Invoices.Currency, Tbl_Invoices.Customer_ID, Tbl"
        "_Invoices.Update_date HAVING (((Tbl_Invoices.Update_date)=#8/20/2024#)); "
    Caption ="Sottomaschera Tbl_Invoices1"
    OnOpen ="[Event Procedure]"
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
            Height =1167
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =114
                    Width =465
                    Height =255
                    ColumnWidth =465
                    Name ="Currency"
                    ControlSource ="Currency"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =114
                            Width =1560
                            Height =255
                            Name ="Currency_Etichetta"
                            Caption ="Currency"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =456
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="Customer_ID"
                    ControlSource ="Customer_ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =456
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
                    Top =798
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="Update_date"
                    ControlSource ="Update_date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =798
                            Width =1560
                            Height =255
                            Name ="Update_date_Etichetta"
                            Caption ="Update_date"
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
CodeBehindForm
' See "SubMaskCurrencies.cls"
