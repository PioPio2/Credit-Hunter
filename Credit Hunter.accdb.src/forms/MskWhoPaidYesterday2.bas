Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =16845
    DatasheetFontHeight =10
    ItemSuffix =50
    Right =16560
    Bottom =11865
    RecSrcDt = Begin
        0xa35a09ab3c3ee640
    End
    RecordSource ="SELECT Tbl_Invoices.Update_date, Tbl_Invoices.*, Tbl_Types.Descripition FROM Tbl"
        "_Invoices LEFT JOIN Tbl_Types ON Tbl_Invoices.Type = Tbl_Types.ID WHERE (((Tbl_I"
        "nvoices.Update_date)=#08/20/2024#)); "
    Caption ="Who paid yesterday"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Tab
            Width =5103
            Height =3402
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =13039
            BackColor =-2147483633
            Name ="SendStatement"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =120
                    Width =2211
                    Height =340
                    FontSize =10
                    FontWeight =700
                    Name ="Testo10"
                    ControlSource ="Customer_code"

                    LayoutCachedLeft =165
                    LayoutCachedTop =120
                    LayoutCachedWidth =2376
                    LayoutCachedHeight =460
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2550
                    Top =120
                    Width =6696
                    Height =325
                    ColumnWidth =3675
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="testo3"
                    ControlSource ="Name"
                    HorizontalAnchor =2

                    LayoutCachedLeft =2550
                    LayoutCachedTop =120
                    LayoutCachedWidth =9246
                    LayoutCachedHeight =445
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9977
                    Top =120
                    Width =1281
                    Height =325
                    ColumnWidth =855
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Country"
                    ControlSource ="Country"
                    HorizontalAnchor =1

                    LayoutCachedLeft =9977
                    LayoutCachedTop =120
                    LayoutCachedWidth =11258
                    LayoutCachedHeight =445
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14853
                    Top =120
                    Height =737
                    TabIndex =3
                    Name ="Button1"
                    Caption ="Send Single Statement"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =14853
                    LayoutCachedTop =120
                    LayoutCachedWidth =16554
                    LayoutCachedHeight =857
                    Overlaps =1
                End
                Begin Tab
                    OverlapFlags =85
                    BackStyle =0
                    Left =120
                    Top =1755
                    Width =16725
                    Height =11145
                    TabIndex =4
                    TabFixedHeight =454
                    Name ="TabCtl48"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =120
                    LayoutCachedTop =1755
                    LayoutCachedWidth =16845
                    LayoutCachedHeight =12900
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =255
                            Top =2340
                            Width =16455
                            Height =10425
                            Name ="Invoices &cleared"
                            EventProcPrefix ="Invoices__cleared"
                            LayoutCachedLeft =255
                            LayoutCachedTop =2340
                            LayoutCachedWidth =16710
                            LayoutCachedHeight =12765
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =255
                                    Top =2340
                                    Width =14400
                                    Height =6555
                                    Name ="Sottomaschera QueryInvoicesClosedInLastDate"
                                    SourceObject ="Form.Sottomaschera QueryInvoicesClosedInLastDate"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_QueryInvoicesClosedInLastDate"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =2340
                                    LayoutCachedWidth =14655
                                    LayoutCachedHeight =8895
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =255
                            Top =2340
                            Width =16455
                            Height =10425
                            Name ="Sheet1"
                            Caption ="Current &Statement"
                            LayoutCachedLeft =255
                            LayoutCachedTop =2340
                            LayoutCachedWidth =16710
                            LayoutCachedHeight =12765
                            Begin
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =215
                                    Left =360
                                    Top =9885
                                    Width =10500
                                    Height =2846
                                    Name ="SottomascheraTblNotes2"
                                    SourceObject ="Form.SottomascheraTblNotes"
                                    LinkChildFields ="CustomerCode"
                                    LinkMasterFields ="Customer_code"
                                    HorizontalAnchor =2
                                    VerticalAnchor =2

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =9885
                                    LayoutCachedWidth =10860
                                    LayoutCachedHeight =12731
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =360
                                            Top =9645
                                            Width =1875
                                            Height =270
                                            Name ="SottomascheraTblNotes2 Label"
                                            Caption ="Notes"
                                            EventProcPrefix ="SottomascheraTblNotes2_Label"
                                            HorizontalAnchor =1
                                            LayoutCachedLeft =360
                                            LayoutCachedTop =9645
                                            LayoutCachedWidth =2235
                                            LayoutCachedHeight =9915
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =345
                                    Top =2592
                                    Width =16125
                                    Height =6645
                                    TabIndex =1
                                    Name ="Maschera1"
                                    SourceObject ="Form.MskContainer"
                                    LinkChildFields ="Customer_code"
                                    LinkMasterFields ="Customer_code"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =345
                                    LayoutCachedTop =2592
                                    LayoutCachedWidth =16470
                                    LayoutCachedHeight =9237
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =11338
                    Top =120
                    Width =3371
                    Height =340
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =255
                    Name ="Testo99"
                    ControlSource ="Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Customer_Status.* FROM Tbl_Customer_Status; "
                    ColumnWidths ="0;0;0;2268"
                    HorizontalAnchor =1

                    LayoutCachedLeft =11338
                    LayoutCachedTop =120
                    LayoutCachedWidth =14709
                    LayoutCachedHeight =460
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =14853
                    Top =1133
                    Height =737
                    TabIndex =6
                    Name ="btnSendAllStatements"
                    Caption ="Send All Statements"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1

                    LayoutCachedLeft =14853
                    LayoutCachedTop =1133
                    LayoutCachedWidth =16554
                    LayoutCachedHeight =1870
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "MskWhoPaidYesterday2.cls"
