Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7591
    DatasheetFontHeight =11
    ItemSuffix =9
    Right =19200
    Bottom =13095
    RecSrcDt = Begin
        0xeb3d1ff1893ae640
    End
    RecordSource ="SELECT Tbl_Users.Name, Tbl_Users.ID FROM Tbl_Users WHERE (((Tbl_Users.Name) Is N"
        "ot Null)) ORDER BY Tbl_Users.Name; "
    Caption ="Cash collected Vs Cash Target"
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =6115
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =2437
                    Top =1360
                    Height =737
                    Name ="Command4"
                    Caption ="Run report"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2437
                    LayoutCachedTop =1360
                    LayoutCachedWidth =4138
                    LayoutCachedHeight =2097
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =120
                    Top =135
                    Width =7471
                    Height =5980
                    Name ="Box5"
                    LayoutCachedLeft =120
                    LayoutCachedTop =135
                    LayoutCachedWidth =7591
                    LayoutCachedHeight =6115
                End
                Begin ComboBox
                    SpecialEffect =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2154
                    Top =457
                    Width =1985
                    Height =330
                    TabIndex =1
                    Name ="Combo6"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Users.Name, Tbl_Users.ID FROM Tbl_Users WHERE (((Tbl_Users.Name) Is N"
                        "ot Null)) ORDER BY Tbl_Users.Name; "
                    ColumnWidths ="2835;0"

                    LayoutCachedLeft =2154
                    LayoutCachedTop =457
                    LayoutCachedWidth =4139
                    LayoutCachedHeight =787
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =170
                            Top =453
                            Width =1635
                            Height =315
                            Name ="Label7"
                            Caption ="Credit Controller"
                            LayoutCachedLeft =170
                            LayoutCachedTop =453
                            LayoutCachedWidth =1805
                            LayoutCachedHeight =768
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCreditControllerSelection.cls"
