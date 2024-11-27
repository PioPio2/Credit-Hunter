Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =16590
    DatasheetFontHeight =10
    ItemSuffix =81
    Right =19200
    Bottom =13065
    RecSrcDt = Begin
        0x6d47ad677babe340
    End
    RecordSource ="SELECT DISTINCTROW Tbl_Customers.*, Tbl_Customers.Name FROM Tbl_Customers ORDER "
        "BY Tbl_Customers.Name; "
    Caption ="Customers master"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
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
            Height =9276
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =120
                    TabIndex =2
                    Name ="Customer_code"
                    ControlSource ="Customer_code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =120
                            Width =1200
                            Height =240
                            Name ="Etichetta1"
                            Caption ="Customer_code"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =1170
                    Width =2490
                    ColumnWidth =3675
                    TabIndex =4
                    Name ="Name"
                    ControlSource ="Tbl_Customers.Name"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =1170
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =1410
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =1170
                            Width =480
                            Height =240
                            Name ="Etichetta5"
                            Caption ="Name"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1170
                            LayoutCachedWidth =615
                            LayoutCachedHeight =1410
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =1510
                    Width =2490
                    ColumnWidth =7470
                    TabIndex =5
                    Name ="Address"
                    ControlSource ="Address"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =1510
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =1750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =1510
                            Width =660
                            Height =240
                            Name ="Etichetta7"
                            Caption ="Address"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1510
                            LayoutCachedWidth =795
                            LayoutCachedHeight =1750
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =1851
                    Width =2490
                    ColumnWidth =3255
                    TabIndex =6
                    Name ="Address2"
                    ControlSource ="Address2"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =1851
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =2091
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =1847
                            Width =1080
                            Height =240
                            Name ="Etichetta9"
                            Caption ="Address line 2"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1847
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =2087
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =2191
                    Width =2490
                    ColumnWidth =4080
                    TabIndex =7
                    Name ="Address3"
                    ControlSource ="Address3"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =2191
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =2431
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =2192
                            Width =1080
                            Height =240
                            Name ="Etichetta11"
                            Caption ="Address line 3"
                            LayoutCachedLeft =135
                            LayoutCachedTop =2192
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =2432
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =2531
                    Width =2490
                    ColumnWidth =4905
                    TabIndex =8
                    Name ="Address4"
                    ControlSource ="Address4"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =2531
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =2771
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =2537
                            Width =1035
                            Height =240
                            Name ="Etichetta13"
                            Caption ="Addressline 4"
                            LayoutCachedLeft =135
                            LayoutCachedTop =2537
                            LayoutCachedWidth =1170
                            LayoutCachedHeight =2777
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =2871
                    Width =2490
                    ColumnWidth =1005
                    TabIndex =9
                    Name ="Address5"
                    ControlSource ="Address5"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =2871
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =3111
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =2867
                            Width =1080
                            Height =240
                            Name ="Etichetta15"
                            Caption ="Address line 5"
                            LayoutCachedLeft =135
                            LayoutCachedTop =2867
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =3107
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1949
                    Top =3211
                    Width =2490
                    ColumnWidth =855
                    TabIndex =10
                    Name ="Country"
                    ControlSource ="Country"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =3211
                    LayoutCachedWidth =4439
                    LayoutCachedHeight =3451
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =3211
                            Width =660
                            Height =240
                            Name ="Etichetta17"
                            Caption ="Country"
                            LayoutCachedLeft =135
                            LayoutCachedTop =3211
                            LayoutCachedWidth =795
                            LayoutCachedHeight =3451
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1949
                    Top =3568
                    Width =4745
                    ColumnWidth =5700
                    TabIndex =11
                    Name ="Email"
                    ControlSource ="Email"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =3568
                    LayoutCachedWidth =6694
                    LayoutCachedHeight =3808
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =3550
                            Width =1168
                            Height =245
                            Name ="Etichetta33"
                            Caption ="Receiver E-mail"
                            LayoutCachedLeft =150
                            LayoutCachedTop =3550
                            LayoutCachedWidth =1318
                            LayoutCachedHeight =3795
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1949
                    Top =4135
                    Width =4745
                    ColumnWidth =6375
                    TabIndex =12
                    Name ="ccEmail"
                    ControlSource ="ccEmail"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =4135
                    LayoutCachedWidth =6694
                    LayoutCachedHeight =4375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =4175
                            Width =1304
                            Height =245
                            Name ="Etichetta37"
                            Caption ="cc external Email"
                            LayoutCachedLeft =150
                            LayoutCachedTop =4175
                            LayoutCachedWidth =1454
                            LayoutCachedHeight =4420
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1949
                    Top =4475
                    TabIndex =13
                    Name ="TotalInsurance"
                    ControlSource ="TotalInsurance"
                    Format ="€#,##0.00;-€#,##0.00"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =4475
                    LayoutCachedWidth =3650
                    LayoutCachedHeight =4715
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =4487
                            Width =1155
                            Height =240
                            Name ="Etichetta49"
                            Caption ="TotalInsurance"
                            LayoutCachedLeft =150
                            LayoutCachedTop =4487
                            LayoutCachedWidth =1305
                            LayoutCachedHeight =4727
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1949
                    Top =4815
                    Width =1686
                    TabIndex =14
                    Name ="StatementForm"
                    ControlSource ="StatementForm"
                    RowSourceType ="Value List"
                    RowSource ="0;No notes;1;Only query;3;Query+note;4;Only notes"
                    ColumnWidths ="0;2268"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =4815
                    LayoutCachedWidth =3635
                    LayoutCachedHeight =5055
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =4854
                            Width =1185
                            Height =240
                            Name ="Etichetta51"
                            Caption ="StatementForm"
                            LayoutCachedLeft =150
                            LayoutCachedTop =4854
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =5094
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1949
                    Top =5212
                    Width =2781
                    TabIndex =1
                    Name ="CasellaCombinata52"
                    ControlSource ="Language"
                    RowSourceType ="Table/Query"
                    RowSource ="Tbl_Languages"
                    ColumnWidths ="0;2268"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =5212
                    LayoutCachedWidth =4730
                    LayoutCachedHeight =5452
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =5214
                            Width =1185
                            Height =240
                            Name ="Etichetta53"
                            Caption ="Language"
                            LayoutCachedLeft =150
                            LayoutCachedTop =5214
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =5454
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1949
                    Top =460
                    Width =2490
                    ColumnWidth =1590
                    TabIndex =3
                    Name ="Credit_controller"
                    ControlSource ="Credit_controller"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_Users.ID, Tbl_Users.UserName, Tbl_Users.Name FROM Tbl_Users; "
                    ColumnWidths ="0;0;2835"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =135
                            Top =460
                            Width =1275
                            Height =240
                            Name ="Etichetta3"
                            Caption ="Credit_controller"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1949
                    Top =3851
                    Width =2441
                    TabIndex =15
                    Name ="CasellaCombinata54"
                    ControlSource ="ContactNames"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =3851
                    LayoutCachedWidth =4390
                    LayoutCachedHeight =4091
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =3881
                            Width =1185
                            Height =240
                            Name ="Etichetta55"
                            Caption ="Contact Name"
                            LayoutCachedLeft =150
                            LayoutCachedTop =3881
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =4121
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =7404
                    Top =6426
                    Width =6804
                    Height =2830
                    Name ="Sottomaschera Tbl_Link_Customer_Internal_Email_Address"
                    SourceObject ="Form.Sottomaschera Tbl_Link_Customer_Internal_Email_Address"
                    LinkChildFields ="CustomerID"
                    LinkMasterFields ="Customer_code"
                    EventProcPrefix ="Sottomaschera_Tbl_Link_Customer_Internal_Email_Address"

                    LayoutCachedLeft =7404
                    LayoutCachedTop =6426
                    LayoutCachedWidth =14208
                    LayoutCachedHeight =9256
                End
                Begin Subform
                    OverlapFlags =85
                    Left =136
                    Top =6432
                    Width =6804
                    Height =2844
                    TabIndex =16
                    Name ="Sottomaschera Tbl_Link_Customer_Internal_Email_Address_Avaialbe"
                    SourceObject ="Form.Sottomaschera Internal_Email_Addresses_Available"
                    EventProcPrefix ="Sottomaschera_Tbl_Link_Customer_Internal_Email_Address_Avaialbe"

                    LayoutCachedLeft =136
                    LayoutCachedTop =6432
                    LayoutCachedWidth =6940
                    LayoutCachedHeight =9276
                End
                Begin Label
                    OverlapFlags =85
                    Left =135
                    Top =6111
                    Width =3329
                    Height =231
                    Name ="Label66"
                    Caption ="Available e-mails"
                    LayoutCachedLeft =135
                    LayoutCachedTop =6111
                    LayoutCachedWidth =3464
                    LayoutCachedHeight =6342
                End
                Begin Label
                    OverlapFlags =85
                    Left =7404
                    Top =6105
                    Width =3274
                    Height =245
                    Name ="Label67"
                    Caption ="Receiver e-mails"
                    LayoutCachedLeft =7404
                    LayoutCachedTop =6105
                    LayoutCachedWidth =10678
                    LayoutCachedHeight =6350
                End
                Begin Line
                    OverlapFlags =85
                    Left =282
                    Top =5470
                    Width =13562
                    Name ="Line68"
                    LayoutCachedLeft =282
                    LayoutCachedTop =5470
                    LayoutCachedWidth =13844
                    LayoutCachedHeight =5470
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6872
                    Top =373
                    Width =3225
                    Height =2460
                    TabIndex =17
                    Name ="Tbl_Types subform"
                    SourceObject ="Form.MskTypesSubform"
                    EventProcPrefix ="Tbl_Types_subform"

                    LayoutCachedLeft =6872
                    LayoutCachedTop =373
                    LayoutCachedWidth =10097
                    LayoutCachedHeight =2833
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6870
                            Top =90
                            Width =2775
                            Height =240
                            Name ="Tbl_Types subform Label"
                            Caption ="Available documents in the statement"
                            EventProcPrefix ="Tbl_Types_subform_Label"
                            LayoutCachedLeft =6870
                            LayoutCachedTop =90
                            LayoutCachedWidth =9645
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =10547
                    Top =373
                    Width =3240
                    Height =2476
                    TabIndex =18
                    Name ="MskDocumentsToBeErasedSubform"
                    SourceObject ="Form.MskDocumentsToBeErasedSubform"
                    LinkChildFields ="CustomerID"
                    LinkMasterFields ="Customer_code"

                    LayoutCachedLeft =10547
                    LayoutCachedTop =373
                    LayoutCachedWidth =13787
                    LayoutCachedHeight =2849
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10545
                            Top =90
                            Width =3120
                            Height =240
                            Name ="MskDocumentsToBeErasedSubform Label"
                            Caption ="Documents to be Deleted in the statement"
                            EventProcPrefix ="MskDocumentsToBeErasedSubform_Label"
                            LayoutCachedLeft =10545
                            LayoutCachedTop =90
                            LayoutCachedWidth =13665
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6885
                    Top =3210
                    TabIndex =19
                    Name ="Check12"
                    ControlSource ="FacturaNumberToBePrinted"

                    LayoutCachedLeft =6885
                    LayoutCachedTop =3210
                    LayoutCachedWidth =7145
                    LayoutCachedHeight =3450
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7110
                            Top =3210
                            Width =4305
                            Height =225
                            Name ="Label13"
                            Caption ="Factura number will be printed in this customer's statement"
                            LayoutCachedLeft =7110
                            LayoutCachedTop =3210
                            LayoutCachedWidth =11415
                            LayoutCachedHeight =3435
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6885
                    Top =3510
                    TabIndex =20
                    Name ="Check72"
                    ControlSource ="PullTicketNumberToBePrinted"

                    LayoutCachedLeft =6885
                    LayoutCachedTop =3510
                    LayoutCachedWidth =7145
                    LayoutCachedHeight =3750
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7110
                            Top =3510
                            Width =4470
                            Height =225
                            Name ="Label73"
                            Caption ="Pull Ticket Number will be printed in this customer's statement"
                            LayoutCachedLeft =7110
                            LayoutCachedTop =3510
                            LayoutCachedWidth =11580
                            LayoutCachedHeight =3735
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6885
                    Top =3810
                    TabIndex =21
                    Name ="Check74"
                    ControlSource ="OriginalInvoiceAmountToBePrinted"

                    LayoutCachedLeft =6885
                    LayoutCachedTop =3810
                    LayoutCachedWidth =7145
                    LayoutCachedHeight =4050
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7110
                            Top =3810
                            Width =5250
                            Height =225
                            Name ="Label75"
                            Caption ="Original Transaction amounts will be printed in this customer's statement"
                            LayoutCachedLeft =7110
                            LayoutCachedTop =3810
                            LayoutCachedWidth =12360
                            LayoutCachedHeight =4035
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1949
                    Top =813
                    Width =2480
                    TabIndex =22
                    Name ="Text79"
                    ControlSource ="RetailOEM"

                    LayoutCachedLeft =1949
                    LayoutCachedTop =813
                    LayoutCachedWidth =4429
                    LayoutCachedHeight =1053
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =150
                            Top =795
                            Width =1168
                            Height =245
                            Name ="Label80"
                            Caption ="Channel"
                            LayoutCachedLeft =150
                            LayoutCachedTop =795
                            LayoutCachedWidth =1318
                            LayoutCachedHeight =1040
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCustomers.cls"
