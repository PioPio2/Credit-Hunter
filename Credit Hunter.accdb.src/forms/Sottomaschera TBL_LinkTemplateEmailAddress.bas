Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2424
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =675
    Top =5970
    Right =10620
    Bottom =7425
    RecSrcDt = Begin
        0xb54fee29fd82e340
    End
    RecordSource ="TBL_LinkTemplateEmailAddress"
    Caption ="Sottomaschera TBL_LinkTemplateEmailAddress"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="IntestazioneMaschera"
        End
        Begin Section
            Height =3558
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin ListBox
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =57
                    Top =369
                    Width =2310
                    Height =1365
                    ColumnWidth =2310
                    Name ="IDTemplate"
                    ControlSource ="IDTemplate"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Templates"
                    ColumnWidths ="0;0;0;0;3402"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =57
                            Top =114
                            Width =1560
                            Height =255
                            Name ="IDTemplate_Etichetta"
                            Caption ="IDTemplate"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =57
                    Top =2079
                    Width =2310
                    Height =1365
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="IDDepartment"
                    ControlSource ="IDDepartment"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_DepartmentNames"
                    ColumnWidths ="0;2268"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =57
                            Top =1824
                            Width =1560
                            Height =255
                            Name ="IDDepartment_Etichetta"
                            Caption ="IDDepartment"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="PièDiPaginaMaschera"
        End
    End
End
