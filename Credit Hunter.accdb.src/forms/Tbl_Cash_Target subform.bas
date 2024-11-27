Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6617
    DatasheetFontHeight =11
    ItemSuffix =6
    Right =13485
    Bottom =13050
    RecSrcDt = Begin
        0x51d471852dcde340
    End
    RecordSource ="Tbl_Cash_Target"
    Caption ="Tbl_Cash_Target subform"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =161
            FontSize =11
            BorderColor =-2147483609
            ForeColor =11830108
            FontName ="Calibri"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =-2147483609
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =-2147483609
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =161
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =14919545
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =161
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =161
            Width =283
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =14919545
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =161
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =14919545
            Name ="FormHeader"
            AutoHeight =1
        End
        Begin Section
            Height =1626
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2865
                    Top =345
                    Width =3660
                    Height =330
                    ColumnWidth =1530
                    Name ="CControllerID"
                    ControlSource ="CControllerID"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2865
                    LayoutCachedTop =345
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =675
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =345
                            Width =2460
                            Height =330
                            Name ="CControllerID_Label"
                            Caption ="CControllerID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =345
                            LayoutCachedWidth =2805
                            LayoutCachedHeight =675
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2865
                    Top =750
                    Width =3660
                    Height =330
                    ColumnWidth =3000
                    TabIndex =1
                    Name ="Channel"
                    ControlSource ="Channel"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2865
                    LayoutCachedTop =750
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =1080
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =750
                            Width =2460
                            Height =330
                            Name ="Channel_Label"
                            Caption ="Channel"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =750
                            LayoutCachedWidth =2805
                            LayoutCachedHeight =1080
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2865
                    Top =1155
                    Width =3660
                    Height =330
                    ColumnWidth =3000
                    TabIndex =2
                    Name ="CashTargetInEUR"
                    ControlSource ="CashTargetInEUR"
                    Format ="€#,##0.00;-€#,##0.00"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2865
                    LayoutCachedTop =1155
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =1485
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1155
                            Width =2460
                            Height =330
                            Name ="CashTargetInEUR_Label"
                            Caption ="CashTargetInEUR"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1155
                            LayoutCachedWidth =2805
                            LayoutCachedHeight =1485
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
        End
    End
End
