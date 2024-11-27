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
    Width =4044
    DatasheetFontHeight =10
    ItemSuffix =28
    Right =16560
    Bottom =11865
    RecSrcDt = Begin
        0xa96216428a3ae640
    End
    RecordSource ="SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invo"
        "ices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=DMax(\"[U"
        "pdate_date]\",\"[Tbl_Invoices]\"))) ORDER BY Tbl_Invoices.Overdue_Date; "
    Caption ="Sottomaschera Tbl_Invoices"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnMouseMove ="[Event Procedure]"
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
            Height =6122
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =450
                    Width =1035
                    Height =255
                    ColumnWidth =1140
                    Name ="Date"
                    ControlSource ="Date"
                    ShowDatePicker =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =450
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =705
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =450
                            Width =1560
                            Height =255
                            Name ="Date_Etichetta"
                            Caption ="Date"
                            LayoutCachedLeft =60
                            LayoutCachedTop =450
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =705
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =792
                    Width =960
                    Height =255
                    ColumnWidth =1770
                    TabIndex =1
                    Name ="Document_Number"
                    ControlSource ="Document_Number"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =792
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1047
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =792
                            Width =1560
                            Height =255
                            Name ="Document_Number_Etichetta"
                            Caption ="Doc. n#"
                            LayoutCachedLeft =60
                            LayoutCachedTop =792
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1047
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =1134
                    Width =975
                    Height =255
                    ColumnWidth =1815
                    TabIndex =2
                    Name ="Customer_reference"
                    ControlSource ="Customer_reference"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1134
                    LayoutCachedWidth =2655
                    LayoutCachedHeight =1389
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1134
                            Width =795
                            Height =255
                            Name ="Customer_reference_Etichetta"
                            Caption ="Reference"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1134
                            LayoutCachedWidth =855
                            LayoutCachedHeight =1389
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =2508
                    Width =2310
                    Height =255
                    ColumnWidth =1095
                    TabIndex =6
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2508
                            Width =1560
                            Height =255
                            Name ="Amount_Etichetta"
                            Caption ="Amount"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =2850
                    Width =1440
                    Height =255
                    ColumnWidth =1140
                    TabIndex =7
                    Name ="Overdue_Date"
                    ControlSource ="Tbl_Invoices.Overdue_Date"
                    ShowDatePicker =0

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2850
                            Width =1560
                            Height =255
                            Name ="Overdue_Date_Etichetta"
                            Caption ="Due date"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1680
                    Top =4200
                    Width =2175
                    ColumnWidth =2744
                    TabIndex =8
                    Name ="Testo20"
                    ControlSource ="Query"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_queries.ID, Tbl_queries.Query, Tbl_queries.Resolution_owner FROM Tbl_"
                        "queries ORDER BY Tbl_queries.Query; "
                    ColumnWidths ="0;2268;2268"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =4200
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =4440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =4200
                            Width =1560
                            Height =255
                            Name ="Etichetta21"
                            Caption ="Query"
                            LayoutCachedLeft =60
                            LayoutCachedTop =4200
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =4455
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1683
                    Top =4563
                    Width =2175
                    Height =255
                    ColumnWidth =4211
                    TabIndex =9
                    Name ="Testo24"
                    ControlSource ="Memo"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =1683
                    LayoutCachedTop =4563
                    LayoutCachedWidth =3858
                    LayoutCachedHeight =4818
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =63
                            Top =4563
                            Width =1560
                            Height =255
                            Name ="Etichetta25"
                            Caption ="Notes"
                            LayoutCachedLeft =63
                            LayoutCachedTop =4563
                            LayoutCachedWidth =1623
                            LayoutCachedHeight =4818
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =3465
                    Width =960
                    Height =255
                    TabIndex =5
                    Name ="Text34"
                    ControlSource ="OriginalAmount"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3465
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =3720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =3465
                            Width =1560
                            Height =255
                            Name ="Label35"
                            Caption ="Original Amount"
                            LayoutCachedTop =3465
                            LayoutCachedWidth =1560
                            LayoutCachedHeight =3720
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =1605
                    Width =975
                    Height =255
                    TabIndex =3
                    Name ="TextSOn#"
                    ControlSource ="SONumber"
                    EventProcPrefix ="TextSOn_"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =1605
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1605
                            Width =1365
                            Height =255
                            Name ="Label39"
                            Caption ="SO number"
                            LayoutCachedTop =1605
                            LayoutCachedWidth =1365
                            LayoutCachedHeight =1860
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
                    ControlSource ="Tbl_Invoices.Currency"

                    LayoutCachedLeft =1677
                    LayoutCachedTop =2166
                    LayoutCachedWidth =3747
                    LayoutCachedHeight =2406
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2166
                            Width =1560
                            Height =255
                            Name ="Type_Etichetta"
                            Caption ="Currency"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =90
                    Width =1035
                    Height =255
                    TabIndex =10
                    Name ="Text26"
                    ControlSource ="ID"
                    ShowDatePicker =0

                    LayoutCachedLeft =1680
                    LayoutCachedTop =90
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =60
                            Top =90
                            Width =1560
                            Height =255
                            Name ="Label27"
                            Caption ="Date"
                            LayoutCachedLeft =60
                            LayoutCachedTop =90
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =345
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
' See "SubMaskTblInvoices2RelIV.cls"
