Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10881
    DatasheetFontHeight =10
    ItemSuffix =13
    Left =611
    Top =5366
    Right =14590
    Bottom =6480
    RecSrcDt = Begin
        0xee11a6677babe340
    End
    RecordSource ="SELECT Tbl_Invoices.Currency, Sum((IIf([overdue_date]>Now(),[amount],0))) AS Cur"
        "r, Tbl_Invoices.Currency, Tbl_Invoices.Customer_ID, Sum((IIf([overdue_date]<=Now"
        "() And [overdue_date]>Now()-31,[amount],0))) AS [1-30], Sum((IIf([overdue_date]<"
        "=Now()-31 And [overdue_date]>Now()-60,[amount],0))) AS [31-60], Sum((IIf([overdu"
        "e_date]<=Now()-61,[amount],0))) AS [60+], Sum((IIf([overdue_date]<=GetNextMonthE"
        "nd(),[amount],0))) AS OverdueMonthEnd FROM Tbl_Invoices WHERE (((Tbl_Invoices.Up"
        "date_date)=Date())) GROUP BY Tbl_Invoices.Currency, Tbl_Invoices.Currency, Tbl_I"
        "nvoices.Customer_ID; "
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin FormHeader
            Height =396
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4280
                    Width =1263
                    Height =231
                    FontWeight =700
                    Name ="Label1"
                    Caption ="1-30 Days:"
                End
                Begin Label
                    OverlapFlags =85
                    Left =122
                    Width =774
                    Height =231
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Currency:"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5720
                    Width =1263
                    Height =231
                    FontWeight =700
                    Name ="Label4"
                    Caption ="31-60 Days:"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7159
                    Width =1263
                    Height =231
                    FontWeight =700
                    Name ="Label6"
                    Caption ="60+Days:"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1182
                    Width =1263
                    Height =230
                    FontWeight =700
                    Name ="Label8"
                    Caption ="Current:"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2839
                    Width =1277
                    Height =231
                    FontWeight =700
                    Name ="Label10"
                    Caption ="Total overdue:"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8939
                    Width =1942
                    Height =231
                    FontWeight =700
                    Name ="Label12"
                    Caption ="Overdue month end"
                End
            End
        End
        Begin Section
            Height =240
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =1
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2622
                    Width =123
                    Height =228
                    Name ="Customer_ID"
                    ControlSource ="Customer_ID"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4279
                    Width =1263
                    TabIndex =1
                    Name ="Amount"
                    ControlSource ="1-30"
                    Format ="Standard"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =109
                    Width =805
                    TabIndex =2
                    Name ="Currency"
                    ControlSource ="Expr1000"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5719
                    Width =1263
                    TabIndex =3
                    Name ="Text3"
                    ControlSource ="31-60"
                    Format ="Standard"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7159
                    Width =1263
                    TabIndex =4
                    Name ="Text5"
                    ControlSource ="60+"
                    Format ="Standard"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1263
                    Width =1267
                    TabIndex =5
                    Name ="Text7"
                    ControlSource ="Curr"
                    Format ="Standard"

                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2835
                    Width =1263
                    TabIndex =6
                    Name ="Text9"
                    ControlSource ="=[1-30]+[31-60]+[60+]"
                    Format ="Standard"

                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    BorderWidth =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8953
                    Width =1889
                    Height =227
                    TabIndex =7
                    Name ="Text11"
                    ControlSource ="OverdueMonthEnd"
                    Format ="Standard"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "MskAging.cls"
