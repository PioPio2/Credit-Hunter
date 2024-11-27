Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4682
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =245
    Top =543
    Right =5217
    Bottom =3097
    RecSrcDt = Begin
        0xcb886ecfb47fe340
    End
    RecordSource ="Tbl_EmailAddresses"
    Caption ="Tbl_EmailAddresses"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =113
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =231
                            Height =231
                            Name ="Label1"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =453
                    Width =2554
                    ColumnWidth =2337
                    TabIndex =1
                    Name ="EmailAddress"
                    ControlSource ="EmailAddress"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =1019
                            Height =231
                            Name ="Label3"
                            Caption ="EmailAddress"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1984
                    Top =793
                    Width =2554
                    Height =1659
                    ColumnWidth =2866
                    TabIndex =2
                    BoundColumn =1
                    Name ="Department"
                    ControlSource ="Department"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_DepartmentNames.DeparmentName, Tbl_DepartmentNames.ID FROM Tbl_Depart"
                        "mentNames; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
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
' See "MskTblEmailAddresses.cls"
