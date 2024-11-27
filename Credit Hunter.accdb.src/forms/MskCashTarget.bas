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
    Width =6465
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =8190
    Bottom =11190
    RecSrcDt = Begin
        0x18fb56c430cde340
    End
    Caption ="Cash Target"
    OnOpen ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin Section
            CanGrow = NotDefault
            Height =4657
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =510
                    Top =637
                    Width =5910
                    Height =4020
                    Name ="MskCashTargetDetails"
                    SourceObject ="Form.MskCashTargetDetails"

                    LayoutCachedLeft =510
                    LayoutCachedTop =637
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =4657
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =240
                            Width =5955
                            Height =315
                            Name ="Label1"
                            Caption ="Cash Targets in EUR currency to be collected within: 30-Sep-2011"
                            LayoutCachedLeft =510
                            LayoutCachedTop =240
                            LayoutCachedWidth =6465
                            LayoutCachedHeight =555
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCashTarget.cls"
