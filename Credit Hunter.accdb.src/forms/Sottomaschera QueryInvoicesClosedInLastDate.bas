Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10318
    DatasheetFontHeight =10
    ItemSuffix =26
    Right =16005
    Bottom =12555
    RecSrcDt = Begin
        0xfb8823b84f3ae640
    End
    RecordSource ="QueryInvoicesClosedInLastDate"
    Caption ="Sottomaschera QueryInvoicesClosedInLastDate"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
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
            Height =3630
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1314
                    Top =798
                    Width =1035
                    Height =255
                    ColumnWidth =1365
                    Name ="Date"
                    ControlSource ="Date"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =798
                            Width =1197
                            Height =255
                            Name ="Date_Etichetta"
                            Caption ="Doc Date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1314
                    Top =1140
                    Width =1800
                    Height =255
                    ColumnWidth =2175
                    TabIndex =1
                    Name ="Document_Number"
                    ControlSource ="Document_Number"

                    LayoutCachedLeft =1314
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3114
                    LayoutCachedHeight =1395
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1140
                            Width =1197
                            Height =255
                            Name ="Document_Number_Etichetta"
                            Caption ="Doc. n#"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1314
                    Top =1482
                    Width =1764
                    Height =255
                    ColumnWidth =2310
                    TabIndex =2
                    Name ="Customer_reference"
                    ControlSource ="Customer_reference"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1482
                            Width =1197
                            Height =255
                            Name ="Customer_reference_Etichetta"
                            Caption ="Customer reference"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1314
                    Top =1824
                    Width =2085
                    Height =255
                    ColumnWidth =1725
                    TabIndex =3
                    Name ="Type"
                    ControlSource ="Descripition"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1824
                            Width =1197
                            Height =255
                            Name ="Type_Etichetta"
                            Caption ="Type"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1317
                    Top =2610
                    Width =1764
                    Height =255
                    ColumnWidth =1650
                    TabIndex =4
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2610
                            Width =1197
                            Height =255
                            Name ="Amount_Etichetta"
                            Caption ="Amount"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1317
                    Top =2190
                    Width =465
                    Height =255
                    ColumnWidth =1335
                    TabIndex =5
                    Name ="Currency"
                    ControlSource ="Currency"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2190
                            Width =1197
                            Height =255
                            Name ="Currency_Etichetta"
                            Caption ="Currency"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1317
                    Top =2985
                    Width =1764
                    Height =255
                    TabIndex =6
                    Name ="Testo22"
                    ControlSource ="Overdue_Date"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2985
                            Width =1197
                            Height =255
                            Name ="Etichetta23"
                            Caption ="Overdue date"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1317
                    Top =3360
                    Width =4134
                    Height =270
                    ColumnWidth =3075
                    TabIndex =7
                    Name ="Testo24"
                    ControlSource ="Tbl_queries.Query"
                    Format ="€#,##0.00;-€#,##0.00"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3360
                            Width =1197
                            Height =255
                            Name ="Etichetta25"
                            Caption ="Query"
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
' See "Sottomaschera QueryInvoicesClosedInLastDate.cls"
