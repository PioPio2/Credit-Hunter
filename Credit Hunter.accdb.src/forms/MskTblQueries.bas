Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4582
    DatasheetFontHeight =10
    ItemSuffix =9
    Right =14910
    Bottom =9930
    RecSrcDt = Begin
        0x08cabc677babe340
    End
    RecordSource ="SELECT Tbl_queries.Query, Tbl_queries.Resolution_owner, Tbl_queries.InvoiceToBeP"
        "aid, Tbl_queries.ID FROM Tbl_queries ORDER BY Tbl_queries.Query; "
    Caption ="Queries"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Section
            Height =2048
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =3120
                    Name ="Query"
                    ControlSource ="Query"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =525
                            Height =240
                            Name ="Label3"
                            Caption ="Query"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =3450
                    TabIndex =1
                    Name ="Resolution_owner"
                    ControlSource ="Resolution_owner"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =1365
                            Height =240
                            Name ="Label5"
                            Caption ="Resolution_owner"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1934
                    Top =1260
                    TabIndex =2
                    Name ="InvoiceToBePaid"
                    ControlSource ="InvoiceToBePaid"

                    LayoutCachedLeft =1934
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2194
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1260
                            Width =1644
                            Height =788
                            Name ="Label6"
                            Caption ="Invoice to be paid ?\015\012Yes=Ticked\015\012No=Unticked"
                            LayoutCachedLeft =120
                            LayoutCachedTop =1260
                            LayoutCachedWidth =1764
                            LayoutCachedHeight =2048
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1922
                    Top =72
                    Width =402
                    TabIndex =3
                    Name ="Text7"
                    ControlSource ="ID"

                    LayoutCachedLeft =1922
                    LayoutCachedTop =72
                    LayoutCachedWidth =2324
                    LayoutCachedHeight =312
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =108
                            Top =72
                            Width =525
                            Height =240
                            Name ="Label8"
                            Caption ="ID"
                            LayoutCachedLeft =108
                            LayoutCachedTop =72
                            LayoutCachedWidth =633
                            LayoutCachedHeight =312
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskTblQueries.cls"
