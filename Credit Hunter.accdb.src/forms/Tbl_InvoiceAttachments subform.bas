Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8892
    DatasheetFontHeight =11
    ItemSuffix =8
    Right =18945
    Bottom =13095
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x2138a4c572e1e340
    End
    RecordSource ="Tbl_InvoiceAttachments"
    Caption ="Tbl_InvoiceAttachments subform"
    OnDelete ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
            Height =2301
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2625
                    Top =345
                    Width =6165
                    Height =600
                    ColumnWidth =7605
                    Name ="AttachName"
                    ControlSource ="AttachName"
                    StatusBarText ="Filename+3 digits number+extension"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2625
                    LayoutCachedTop =345
                    LayoutCachedWidth =8790
                    LayoutCachedHeight =945
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
                            Width =2223
                            Height =600
                            Name ="AttachName_Label"
                            Caption ="Attachment"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =345
                            LayoutCachedWidth =2568
                            LayoutCachedHeight =945
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2625
                    Top =1020
                    Width =6165
                    Height =330
                    ColumnWidth =3660
                    TabIndex =1
                    Name ="DocumentID"
                    ControlSource ="DocumentID"
                    StatusBarText ="n#+Date+Type+Reference"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2625
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8790
                    LayoutCachedHeight =1350
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
                            Top =1020
                            Width =2223
                            Height =330
                            Name ="DocumentID_Label"
                            Caption ="DocumentID"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2568
                            LayoutCachedHeight =1350
                            RowStart =1
                            RowEnd =1
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
CodeBehindForm
' See "Tbl_InvoiceAttachments subform.cls"
