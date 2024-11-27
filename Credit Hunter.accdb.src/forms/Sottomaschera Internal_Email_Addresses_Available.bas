Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    Width =3686
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =1032
    Top =3614
    Right =7512
    Bottom =6127
    RecSrcDt = Begin
        0x1b9cc3677babe340
    End
    RecordSource ="SELECT Tbl_EmailAddresses.Department, Tbl_EmailAddresses.EmailAddress, Tbl_Email"
        "Addresses.ID, Tbl_Link_Customer_Internal_Email_Address.InternalEmailAddressID, T"
        "bl_DepartmentNames.DeparmentName FROM (Tbl_EmailAddresses LEFT JOIN Tbl_Link_Cus"
        "tomer_Internal_Email_Address ON Tbl_EmailAddresses.ID=Tbl_Link_Customer_Internal"
        "_Email_Address.InternalEmailAddressID) LEFT JOIN Tbl_DepartmentNames ON Tbl_Emai"
        "lAddresses.Department=Tbl_DepartmentNames.ID WHERE (((Tbl_Link_Customer_Internal"
        "_Email_Address.CustomerID) Is Null Or (Tbl_Link_Customer_Internal_Email_Address."
        "CustomerID)<>32246)) GROUP BY Tbl_EmailAddresses.Department, Tbl_EmailAddresses."
        "EmailAddress, Tbl_EmailAddresses.ID, Tbl_Link_Customer_Internal_Email_Address.In"
        "ternalEmailAddressID, Tbl_DepartmentNames.DeparmentName ORDER BY Tbl_EmailAddres"
        "ses.Department, Tbl_EmailAddresses.EmailAddress; "
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
            Height =2834
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1670
                    Top =1195
                    Width =1443
                    Height =255
                    ColumnWidth =4469
                    Name ="Testo4"
                    ControlSource ="EmailAddress"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =109
                            Top =1202
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
                    Left =1674
                    Top =1522
                    Width =1838
                    Height =255
                    ColumnWidth =1644
                    TabIndex =1
                    Name ="Testo6"
                    ControlSource ="DeparmentName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =1522
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
            Height =1303
            BackColor =-2147483633
            Name ="PièDiPaginaMaschera"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1674
                    Top =109
                    Width =1443
                    Height =255
                    Name ="Text10"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =54
                            Top =109
                            Width =1560
                            Height =255
                            Name ="Label11"
                            Caption ="E-mail address"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Sottomaschera Internal_Email_Addresses_Available.cls"
