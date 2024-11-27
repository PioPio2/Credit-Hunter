Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =161
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =510
    Top =630
    Right =6405
    Bottom =4635
    RecSrcDt = Begin
        0x966f1af1893ae640
    End
    RecordSource ="SELECT Tbl_Cash_Target.CControllerID, Tbl_Cash_Target.Channel, Tbl_Cash_Target.C"
        "ashTargetInEUR, Tbl_Users.Name FROM Tbl_Cash_Target LEFT JOIN Tbl_Users ON Tbl_C"
        "ash_Target.CControllerID = Tbl_Users.ID; "
    DatasheetFontName ="Calibri"
    OnKeyDown ="[Event Procedure]"
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
        Begin Section
            Height =5952
            BackColor =-2147483633
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1133
                    Top =963
                    Height =315
                    ColumnWidth =2100
                    Name ="Text0"
                    ControlSource ="Name"

                    LayoutCachedLeft =1133
                    LayoutCachedTop =963
                    LayoutCachedWidth =2834
                    LayoutCachedHeight =1278
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =960
                            Width =1605
                            Height =315
                            Name ="Label1"
                            Caption ="Credit controller"
                            LayoutCachedTop =960
                            LayoutCachedWidth =1605
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1133
                    Top =1398
                    Height =315
                    TabIndex =1
                    Name ="Text2"
                    ControlSource ="Channel"

                    LayoutCachedLeft =1133
                    LayoutCachedTop =1398
                    LayoutCachedWidth =2834
                    LayoutCachedHeight =1713
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1395
                            Width =840
                            Height =315
                            Name ="Label3"
                            Caption ="Channel"
                            LayoutCachedTop =1395
                            LayoutCachedWidth =840
                            LayoutCachedHeight =1710
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1133
                    Top =1833
                    Height =315
                    ColumnWidth =1650
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="CashTargetInEUR"

                    LayoutCachedLeft =1133
                    LayoutCachedTop =1833
                    LayoutCachedWidth =2834
                    LayoutCachedHeight =2148
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1830
                            Width =1110
                            Height =315
                            Name ="Label5"
                            Caption ="Cash target"
                            LayoutCachedTop =1830
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =2145
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashTargetDetails.cls"
