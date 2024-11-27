Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5159
    RowHeight =272
    DatasheetFontHeight =10
    ItemSuffix =46
    Right =16560
    Bottom =11865
    RecSrcDt = Begin
        0x091d12428a3ae640
    End
    RecordSource ="SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invo"
        "ices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=DMax(\"[U"
        "pdate_date]\",\"[Tbl_Invoices]\"))) ORDER BY Tbl_Invoices.Overdue_Date; "
    Caption ="Sottomaschera Tbl_Invoices"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnMouseMove ="[Event Procedure]"
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
            Height =5782
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =735
                    Width =1035
                    Height =255
                    ColumnWidth =1245
                    ColumnOrder =0
                    Name ="Date"
                    ControlSource ="Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =735
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =990
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =735
                            Width =1560
                            Height =255
                            Name ="Date_Etichetta"
                            Caption ="Date"
                            LayoutCachedLeft =56
                            LayoutCachedTop =735
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =990
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1077
                    Width =960
                    Height =255
                    ColumnWidth =2010
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Document_Number"
                    ControlSource ="Document_Number"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1077
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1332
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =1077
                            Width =1560
                            Height =255
                            Name ="Document_Number_Etichetta"
                            Caption ="Doc. n#"
                            LayoutCachedLeft =56
                            LayoutCachedTop =1077
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =1332
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1419
                    Width =975
                    Height =255
                    ColumnWidth =1815
                    TabIndex =2
                    Name ="Customer_reference"
                    ControlSource ="Customer_reference"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1419
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =1674
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =1419
                            Width =1365
                            Height =255
                            Name ="Customer_reference_Etichetta"
                            Caption ="Reference"
                            LayoutCachedLeft =56
                            LayoutCachedTop =1419
                            LayoutCachedWidth =1421
                            LayoutCachedHeight =1674
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2775
                    Width =2310
                    Height =255
                    ColumnWidth =1515
                    TabIndex =7
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2775
                    LayoutCachedWidth =3990
                    LayoutCachedHeight =3030
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =2775
                            Width =1560
                            Height =255
                            Name ="Amount_Etichetta"
                            Caption ="Amount"
                            LayoutCachedLeft =56
                            LayoutCachedTop =2775
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =3030
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3117
                    Width =1440
                    Height =255
                    ColumnWidth =1140
                    TabIndex =8
                    Name ="Overdue_Date"
                    ControlSource ="Tbl_Invoices.Overdue_Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =3117
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3372
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =3117
                            Width =1560
                            Height =255
                            Name ="Overdue_Date_Etichetta"
                            Caption ="Due date"
                            LayoutCachedLeft =56
                            LayoutCachedTop =3117
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =3372
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1680
                    Top =3834
                    Width =2175
                    ColumnWidth =2744
                    TabIndex =9
                    Name ="Testo20"
                    ControlSource ="Query"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_queries.ID, Tbl_queries.Query, Tbl_queries.Resolution_owner, Tbl_quer"
                        "ies.ToFillChargebackFile FROM Tbl_queries ORDER BY Tbl_queries.Query; "
                    ColumnWidths ="0;2268;2268"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =3834
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =4074
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =3834
                            Width =1560
                            Height =255
                            Name ="Etichetta21"
                            Caption ="Query"
                            LayoutCachedLeft =56
                            LayoutCachedTop =3834
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =4089
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1683
                    Top =4197
                    Width =2175
                    Height =255
                    ColumnWidth =4211
                    TabIndex =10
                    Name ="Testo24"
                    ControlSource ="Memo"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =1683
                    LayoutCachedTop =4197
                    LayoutCachedWidth =3858
                    LayoutCachedHeight =4452
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =4197
                            Width =1560
                            Height =255
                            Name ="Etichetta25"
                            Caption ="Notes"
                            LayoutCachedLeft =56
                            LayoutCachedTop =4197
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =4452
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1870
                    Top =3499
                    Width =960
                    Height =255
                    ColumnWidth =1417
                    TabIndex =6
                    Name ="Text34"
                    ControlSource ="OriginalAmount"

                    LayoutCachedLeft =1870
                    LayoutCachedTop =3499
                    LayoutCachedWidth =2830
                    LayoutCachedHeight =3754
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =3499
                            Width =1560
                            Height =255
                            Name ="Label35"
                            Caption ="Original Amount"
                            LayoutCachedLeft =56
                            LayoutCachedTop =3499
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =3754
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1794
                    Width =975
                    Height =255
                    ColumnWidth =2040
                    TabIndex =3
                    Name ="TextSOn#"
                    ControlSource ="SONumber"
                    EventProcPrefix ="TextSOn_"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1794
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =2049
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =1794
                            Width =1365
                            Height =255
                            Name ="Label39"
                            Caption ="SO number"
                            LayoutCachedLeft =56
                            LayoutCachedTop =1794
                            LayoutCachedWidth =1421
                            LayoutCachedHeight =2049
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1677
                    Top =2166
                    Width =2070
                    TabIndex =4
                    Name ="Type"
                    ControlSource ="Type"

                    LayoutCachedLeft =1677
                    LayoutCachedTop =2166
                    LayoutCachedWidth =3747
                    LayoutCachedHeight =2406
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =2166
                            Width =1560
                            Height =255
                            Name ="Type_Etichetta"
                            Caption ="Type"
                            LayoutCachedLeft =56
                            LayoutCachedTop =2166
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =2421
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1700
                    Top =283
                    Width =450
                    Height =255
                    TabIndex =11
                    Name ="InvoiceID"
                    ControlSource ="ID"
                    ShowDatePicker =0

                    LayoutCachedLeft =1700
                    LayoutCachedTop =283
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =538
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =56
                            Top =226
                            Width =1560
                            Height =255
                            Name ="Label43"
                            Caption ="Date"
                            LayoutCachedLeft =56
                            LayoutCachedTop =226
                            LayoutCachedWidth =1616
                            LayoutCachedHeight =481
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2607
                    Top =2494
                    Width =2070
                    TabIndex =5
                    Name ="LongDescription"
                    ControlSource ="Tbl_Invoices.Currency"

                    LayoutCachedLeft =2607
                    LayoutCachedTop =2494
                    LayoutCachedWidth =4677
                    LayoutCachedHeight =2734
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =2475
                            Width =2145
                            Height =255
                            Name ="Label45"
                            Caption ="Currency"
                            LayoutCachedLeft =56
                            LayoutCachedTop =2475
                            LayoutCachedWidth =2201
                            LayoutCachedHeight =2730
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
' See "SubMaskTblInvoices2Rel.cls"
