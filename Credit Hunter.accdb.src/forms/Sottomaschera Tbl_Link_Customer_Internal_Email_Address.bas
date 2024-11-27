Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3344
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =8137
    Top =4945
    Right =14808
    Bottom =7743
    RecSrcDt = Begin
        0x8dadc9677babe340
    End
    RecordSource ="SELECT Tbl_EmailAddresses.Department, Tbl_Link_Customer_Internal_Email_Address.C"
        "ustomerID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID, Tbl_"
        "EmailAddresses.EmailAddress, Tbl_DepartmentNames.DeparmentName FROM (Tbl_Link_Cu"
        "stomer_Internal_Email_Address INNER JOIN Tbl_EmailAddresses ON Tbl_Link_Customer"
        "_Internal_Email_Address.InternalEmailAddressID=Tbl_EmailAddresses.ID) INNER JOIN"
        " Tbl_DepartmentNames ON Tbl_EmailAddresses.Department=Tbl_DepartmentNames.ID ORD"
        "ER BY Tbl_EmailAddresses.Department, Tbl_EmailAddresses.EmailAddress; "
    Caption ="Sottomaschera Tbl_Link_Customer_Internal_Email_Address"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
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
            Height =1836
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1677
                    Top =831
                    Width =1552
                    Height =255
                    ColumnWidth =4171
                    Name ="Testo4"
                    ControlSource ="EmailAddress"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =831
                            Width =1560
                            Height =255
                            Name ="Etichetta5"
                            Caption ="E-mail address"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1620
                    Top =1223
                    Width =1498
                    Height =255
                    ColumnWidth =1848
                    TabIndex =1
                    Name ="Testo6"
                    ControlSource ="DeparmentName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1223
                            Width =1560
                            Height =255
                            Name ="Etichetta7"
                            Caption ="Team"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1814
            BackColor =-2147483633
            Name ="PièDiPaginaMaschera"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =584
                    Width =1552
                    Height =255
                    Name ="Text12"
                    ControlSource ="InternalEmailAddressID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =584
                            Width =1560
                            Height =255
                            Name ="Label13"
                            Caption ="E-mail address"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =1141
                    Width =1498
                    Height =255
                    TabIndex =1
                    Name ="Text10"
                    ControlSource ="Department"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =1141
                            Width =1560
                            Height =255
                            Name ="Label11"
                            Caption ="Team"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Sottomaschera Tbl_Link_Customer_Internal_Email_Address.cls"
