Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15825
    DatasheetFontHeight =10
    ItemSuffix =41
    Left =330
    Top =1665
    Right =20475
    Bottom =8280
    RecSrcDt = Begin
        0x7dcae20c4f3ae640
    End
    RecordSource ="Tbl_Customers"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
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
            Height =6576
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Tab
                    TabStop = NotDefault
                    MultiRow = NotDefault
                    OverlapFlags =85
                    Width =15825
                    Height =4035
                    Name ="TabCtl87"
                    OnChange ="[Event Procedure]"
                    HorizontalAnchor =2

                    LayoutCachedWidth =15825
                    LayoutCachedHeight =4035
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =90
                            Top =390
                            Width =15606
                            Height =3510
                            Name ="Pagina88"
                            Caption ="EUR"
                            LayoutCachedLeft =90
                            LayoutCachedTop =390
                            LayoutCachedWidth =15696
                            LayoutCachedHeight =3900
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =90
                                    Top =390
                                    Width =15606
                                    Height =3402
                                    Name ="Sottomaschera Tbl_Invoices"
                                    SourceObject ="Form.SubMaskTblInvoices2Rel"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_Tbl_Invoices"
                                    HorizontalAnchor =2

                                    LayoutCachedLeft =90
                                    LayoutCachedTop =390
                                    LayoutCachedWidth =15696
                                    LayoutCachedHeight =3792
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =109
                            Top =405
                            Width =15581
                            Height =3495
                            Name ="Pagina89"
                            Caption ="USD"
                            LayoutCachedLeft =109
                            LayoutCachedTop =405
                            LayoutCachedWidth =15690
                            LayoutCachedHeight =3900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =109
                                    Top =405
                                    Width =11340
                                    Height =3402
                                    Name ="SubMaskTblInvoices2RelII"
                                    SourceObject ="Form.SubMaskTblInvoices2RelII"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"

                                    LayoutCachedLeft =109
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =11449
                                    LayoutCachedHeight =3807
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =109
                            Top =405
                            Width =15581
                            Height =3495
                            Name ="Pagina90"
                            LayoutCachedLeft =109
                            LayoutCachedTop =405
                            LayoutCachedWidth =15690
                            LayoutCachedHeight =3900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =109
                                    Top =405
                                    Width =11340
                                    Height =3402
                                    Name ="SubMaskTblInvoices2RelIII"
                                    SourceObject ="Form.SubMaskTblInvoices2RelIII"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"

                                    LayoutCachedLeft =109
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =11449
                                    LayoutCachedHeight =3807
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =109
                            Top =405
                            Width =15581
                            Height =3495
                            Name ="Pagina91"
                            LayoutCachedLeft =109
                            LayoutCachedTop =405
                            LayoutCachedWidth =15690
                            LayoutCachedHeight =3900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =109
                                    Top =405
                                    Width =11340
                                    Height =3402
                                    Name ="SubMaskTblInvoices2RelIV"
                                    SourceObject ="Form.SubMaskTblInvoices2RelIV"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"

                                    LayoutCachedLeft =109
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =11449
                                    LayoutCachedHeight =3807
                                End
                            End
                        End
                        Begin Page
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =109
                            Top =405
                            Width =15581
                            Height =3495
                            Name ="Pagina92"
                            LayoutCachedLeft =109
                            LayoutCachedTop =405
                            LayoutCachedWidth =15690
                            LayoutCachedHeight =3900
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =109
                                    Top =405
                                    Width =11340
                                    Height =3402
                                    Name ="SubMaskTblInvoices2RelV"
                                    SourceObject ="Form.SubMaskTblInvoices2RelV"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"

                                    LayoutCachedLeft =109
                                    LayoutCachedTop =405
                                    LayoutCachedWidth =11449
                                    LayoutCachedHeight =3807
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =4649
                    Top =4535
                    Width =3005
                    Height =189
                    Name ="Etichetta9"
                    Caption ="Overdue 1-30 days as of today (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =4535
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =4724
                End
                Begin Label
                    OverlapFlags =85
                    Left =4649
                    Top =4780
                    Width =3005
                    Height =216
                    Name ="Etichetta10"
                    Caption ="Overdue 31-60 days as of today (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =4780
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =4996
                End
                Begin Label
                    OverlapFlags =85
                    Left =4649
                    Top =5040
                    Width =3005
                    Height =210
                    Name ="Etichetta11"
                    Caption ="Overdue over 60 as of today (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =5040
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =5250
                End
                Begin Label
                    OverlapFlags =85
                    Left =4649
                    Top =4305
                    Width =3005
                    Height =188
                    Name ="Etichetta12"
                    Caption ="Total overdue as of today (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =4305
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =4493
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7860
                    Top =4535
                    Width =1304
                    Height =189
                    Name ="Etichetta13"
                    Caption ="0.00"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =4535
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =4724
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7860
                    Top =4780
                    Width =1304
                    Height =216
                    Name ="Etichetta14"
                    Caption ="0.00"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =4780
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =4996
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7860
                    Top =5036
                    Width =1304
                    Height =245
                    Name ="Etichetta15"
                    Caption ="-6,295.00"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =5036
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =5281
                End
                Begin Label
                    OverlapFlags =85
                    Left =4649
                    Top =5297
                    Width =3005
                    Height =240
                    FontWeight =700
                    Name ="Etichetta16"
                    Caption ="Total overdue on month end (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =5297
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =5537
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4649
                    Top =5565
                    Width =3005
                    Height =517
                    FontWeight =700
                    BackColor =255
                    Name ="Label19"
                    Caption ="Overdue over 90 days (check insurance obligations) (EUR)"
                    LayoutCachedLeft =4649
                    LayoutCachedTop =5565
                    LayoutCachedWidth =7654
                    LayoutCachedHeight =6082
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =7860
                    Top =5569
                    Width =1304
                    Height =245
                    FontWeight =700
                    BackColor =255
                    Name ="Label20"
                    Caption ="0.00"
                    LayoutCachedLeft =7860
                    LayoutCachedTop =5569
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =5814
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7860
                    Top =4305
                    Width =1304
                    Height =189
                    TabIndex =1
                    Name ="Etichetta8"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =4305
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =4494
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7860
                    Top =5297
                    Width =1304
                    FontWeight =700
                    TabIndex =2
                    Name ="Etichetta17"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =5297
                    LayoutCachedWidth =9164
                    LayoutCachedHeight =5537
                End
                Begin Subform
                    OverlapFlags =87
                    Left =120
                    Top =4335
                    Width =4366
                    Height =2025
                    TabIndex =3
                    Name ="MskLast5Payments"
                    SourceObject ="Form.MskLast5Payments"
                    LinkChildFields ="CustomerID"
                    LinkMasterFields ="Customer_code"

                    LayoutCachedLeft =120
                    LayoutCachedTop =4335
                    LayoutCachedWidth =4486
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =4095
                            Width =3450
                            Height =240
                            Name ="Label31"
                            Caption ="Payments received in the last 365 days"
                            LayoutCachedLeft =120
                            LayoutCachedTop =4095
                            LayoutCachedWidth =3570
                            LayoutCachedHeight =4335
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskContainer.cls"
