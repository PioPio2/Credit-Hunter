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
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4044
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =210
    Top =1845
    Right =14340
    Bottom =4245
    RecSrcDt = Begin
        0x0b5005df3940e340
    End
    RecordSource ="SELECT Tbl_Invoices.*, Tbl_Invoices.Update_date, Tbl_Invoices.Currency, Tbl_Invo"
        "ices.Overdue_Date FROM Tbl_Invoices WHERE (((Tbl_Invoices.Update_date)=Date())) "
        "ORDER BY Tbl_Invoices.Currency, Tbl_Invoices.Overdue_Date; "
    Caption ="Sottomaschera Tbl_Invoices"
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
            Height =4648
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1140
                    Width =1035
                    Height =255
                    ColumnWidth =1140
                    Name ="Date"
                    ControlSource ="Date"
                    ShowDatePicker =0

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
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1482
                    Width =960
                    Height =255
                    ColumnWidth =2051
                    TabIndex =1
                    Name ="Document_Number"
                    ControlSource ="Document_Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1482
                            Width =1560
                            Height =255
                            Name ="Document_Number_Etichetta"
                            Caption ="Document number"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =1824
                    Width =975
                    Height =255
                    ColumnWidth =1815
                    TabIndex =2
                    Name ="Customer_reference"
                    ControlSource ="Customer_reference"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1824
                            Width =795
                            Height =255
                            Name ="Customer_reference_Etichetta"
                            Caption ="Customer reference"
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
                    ColumnWidth =1410
                    TabIndex =4
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
                    TabIndex =5
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
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =3192
                    Width =810
                    Height =255
                    ColumnWidth =1209
                    TabIndex =6
                    Name ="Currency"
                    ControlSource ="Tbl_Invoices.Currency"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =3192
                            Width =1560
                            Height =255
                            Name ="Currency_Etichetta"
                            Caption ="Currency"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1677
                    Top =2166
                    Width =2070
                    ColumnWidth =1290
                    TabIndex =3
                    Name ="Type"
                    ControlSource ="Type"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Types"
                    ColumnWidths ="0;1701;0"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2166
                            Width =1560
                            Height =255
                            Name ="Type_Etichetta"
                            Caption ="Type"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1677
                    Top =3567
                    Width =2175
                    ColumnWidth =2744
                    TabIndex =7
                    Name ="Testo20"
                    ControlSource ="Query"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_queries.ID, Tbl_queries.Query, Tbl_queries.Resolution_owner FROM Tbl_"
                        "queries ORDER BY Tbl_queries.Query; "
                    ColumnWidths ="0;2268;2268"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =3567
                            Width =1560
                            Height =255
                            Name ="Etichetta21"
                            Caption ="Query"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =3930
                    Width =2175
                    Height =255
                    ColumnWidth =4211
                    TabIndex =8
                    Name ="Testo24"
                    ControlSource ="Memo"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3930
                            Width =1560
                            Height =255
                            Name ="Etichetta25"
                            Caption ="Notes"
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
' See "SottomascheraTbl_Invoices.cls"
