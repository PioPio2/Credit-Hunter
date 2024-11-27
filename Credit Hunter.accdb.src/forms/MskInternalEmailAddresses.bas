Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5442
    DatasheetFontHeight =10
    ItemSuffix =6
    Right =10164
    Bottom =10080
    RecSrcDt = Begin
        0xbc82a9058a3ae640
    End
    RecordSource ="SELECT Tbl_EmailAddresses.EmailAddress, Tbl_EmailAddresses.Department, Tbl_Email"
        "Addresses.ID FROM Tbl_EmailAddresses ORDER BY Tbl_EmailAddresses.EmailAddress; "
    Caption ="Internal e-mail addresses"
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2572
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1417
                    Top =453
                    Width =3568
                    Height =227
                    ColumnWidth =2337
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =109
                            Top =448
                            Width =1019
                            Height =231
                            Name ="Label3"
                            Caption ="Email address"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1417
                    Top =793
                    Width =3568
                    Height =1646
                    ColumnWidth =2866
                    TabIndex =1
                    Name ="Department"
                    ControlSource ="Department"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_DepartmentNames.ID, Tbl_DepartmentNames.DeparmentName FROM Tbl_Depart"
                        "mentNames; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =109
                            Top =788
                            Width =965
                            Height =231
                            Name ="Label5"
                            Caption ="Department"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskInternalEmailAddresses.cls"
