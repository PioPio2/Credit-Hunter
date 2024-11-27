Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13895
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =2265
    Top =15
    Right =21165
    Bottom =9195
    RecSrcDt = Begin
        0x987302a01755e340
    End
    RecordSource ="Tbl_Templates"
    Caption ="Templates"
    DatasheetFontName ="Arial"
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7880
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =4949
                    Top =460
                    Width =4908
                    Height =293
                    Name ="Language"
                    ControlSource ="Language"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Tbl_Languages].[ID], [Tbl_Languages].[Language] FROM Tbl_Languages ORDER"
                        " BY [Language]; "
                    ColumnWidths ="0;1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3135
                            Top =460
                            Width =780
                            Height =240
                            Name ="Etichetta3"
                            Caption ="Language"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4949
                    Top =1365
                    Width =8910
                    Height =3900
                    TabIndex =1
                    Name ="Text"
                    ControlSource ="Text"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3135
                            Top =1365
                            Width =405
                            Height =240
                            Name ="Etichetta5"
                            Caption ="Text"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11397
                    Top =480
                    Width =2436
                    TabIndex =3
                    Name ="Testo17"
                    ControlSource ="TemplateName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10035
                            Top =480
                            Width =1200
                            Height =240
                            Name ="Etichetta18"
                            Caption ="Template name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4949
                    Top =810
                    Width =8919
                    Height =278
                    TabIndex =4
                    Name ="Testo20"
                    ControlSource ="Subject"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3135
                            Top =810
                            Width =1584
                            Height =240
                            Name ="Etichetta21"
                            Caption ="Email Subject"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Top =1712
                    Width =4361
                    Height =3437
                    Name ="Etichetta8"
                    Caption ="Available data:\015\012\015\012«1» = Contact name\015\012«2» = Company name\015\012"
                        "«3» = Logitech fiscal month end date\015\012«4» = Total overdue as if fiscal mon"
                        "th end date\015\012«5» = Total overdue up to date\015\012«6» = Customer Credit L"
                        "imit\015\012«7» = Status (Blank, Warning, On hold, Collection mandate)\015\012«1"
                        "0» = Credit Controller signature\015\012«11» = Customer ID\015\012«12»= Today's "
                        "date\015\012«13»= Credit controller's name\015\012«14»= Main recipient order rel"
                        "ease's name\015\012«15» = Last day payment date\015\012«=ALT+174\015\012»=ALT+17"
                        "5"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11168
                    Top =113
                    Width =1595
                    TabIndex =2
                    Name ="Step"
                    ControlSource ="Step"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Customer_Status.Step, Tbl_Customer_Status.Description FROM Tbl_Custom"
                        "er_Status; "
                    ColumnWidths ="567;2268"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10601
                            Top =113
                            Width =405
                            Height =240
                            Name ="Etichetta7"
                            Caption ="Step"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =396
                    Top =5965
                    Width =9975
                    Height =1740
                    TabIndex =5
                    Name ="Sottomaschera TBL_LinkTemplateEmailAddress"
                    SourceObject ="Form.Sottomaschera TBL_LinkTemplateEmailAddress"
                    LinkChildFields ="IDTemplate"
                    LinkMasterFields ="ID"
                    EventProcPrefix ="Sottomaschera_TBL_LinkTemplateEmailAddress"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1362
                    Width =2436
                    TabIndex =6
                    Name ="Testo25"
                    ControlSource ="ID"

                End
            End
        End
    End
End
