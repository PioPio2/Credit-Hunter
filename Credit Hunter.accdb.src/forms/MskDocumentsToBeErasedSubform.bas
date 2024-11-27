Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4487
    RowHeight =345
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =10545
    Top =375
    Right =13755
    Bottom =2820
    RecSrcDt = Begin
        0x2bc927f1893ae640
    End
    RecordSource ="SELECT Tbl_Types.Descripition, Tbl_DocumentsToBeErased.CustomerID, Tbl_Documents"
        "ToBeErased.DocumentType FROM Tbl_DocumentsToBeErased INNER JOIN Tbl_Types ON Tbl"
        "_DocumentsToBeErased.DocumentType = Tbl_Types.ID GROUP BY Tbl_Types.Descripition"
        ", Tbl_DocumentsToBeErased.CustomerID, Tbl_DocumentsToBeErased.DocumentType; "
    Caption ="MskDocumentsToBeErasedSubform"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =163
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
            TextFontCharSet =163
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
            TextFontCharSet =163
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =-2147483609
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =163
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
            TextFontCharSet =163
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
            TextFontCharSet =163
            Width =283
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =14919545
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =163
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
            Height =2324
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2865
                    Top =1200
                    Width =1530
                    Height =330
                    ColumnWidth =2685
                    Name ="Text4"
                    ControlSource ="Descripition"
                    OnDblClick ="[Event Procedure]"
                    GroupTable =2
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2865
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =1530
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1200
                            Width =2460
                            Height =330
                            Name ="Label5"
                            Caption ="Document"
                            GroupTable =2
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2805
                            LayoutCachedHeight =1530
                            LayoutGroup =1
                            GroupTable =2
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =840
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "MskDocumentsToBeErasedSubform.cls"
