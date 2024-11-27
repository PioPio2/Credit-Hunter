Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15000
    DatasheetFontHeight =10
    ItemSuffix =87
    Right =15351
    Bottom =9659
    RecSrcDt = Begin
        0x5b7b29584141e340
    End
    RecordSource ="SELECT * FROM Tbl_Customers WHERE Tbl_Customers.Name LIKE '*loading*'; "
    Caption ="!!!ATTENTION !!! Last update: 30/11/2007 09:16:21"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
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
        Begin Tab
            Width =5103
            Height =3402
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9810
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Width =15000
                    Height =9810
                    Name ="TabCtl48"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =122
                            Top =380
                            Width =14753
                            Height =9306
                            Name ="&Main"
                            EventProcPrefix ="Ctl_Main"
                            Begin
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =215
                                    Top =486
                                    Width =926
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    Name ="Testo1"
                                    ControlSource ="Customer_code"

                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1292
                                    Top =486
                                    Width =5726
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="Testo3"
                                    ControlSource ="Name"

                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =10544
                                    Top =486
                                    Width =1181
                                    Height =345
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="Testo5"
                                    ControlSource ="Credit_controller"

                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =7188
                                    Top =486
                                    Width =3116
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="Testo18"
                                    ControlSource ="Country"

                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =12570
                                    Top =3623
                                    Height =737
                                    TabIndex =4
                                    Name ="Comando21"
                                    Caption ="Proceed"
                                    OnClick ="[Event Procedure]"

                                End
                                Begin OptionGroup
                                    SpecialEffect =1
                                    OverlapFlags =215
                                    Left =12135
                                    Top =2348
                                    Width =2547
                                    Height =2536
                                    TabIndex =5
                                    Name ="Cornice34"
                                    DefaultValue ="1"

                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =215
                                            Left =12402
                                            Top =2235
                                            Width =900
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label35"
                                            Caption ="Statements"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =215
                                            AccessKey =67
                                            Left =12285
                                            Top =2723
                                            OptionValue =1
                                            Name ="Option37"
                                            UnicodeAccessKey =99

                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =12612
                                                    Top =2723
                                                    Width =915
                                                    Height =240
                                                    Name ="Label38"
                                                    Caption ="Only &create"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =215
                                            AccessKey =67
                                            Left =12286
                                            Top =3145
                                            OptionValue =2
                                            Name ="Option39"
                                            UnicodeAccessKey =67

                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =12612
                                                    Top =3145
                                                    Width =1905
                                                    Height =240
                                                    Name ="Label40"
                                                    Caption ="&Create and send email"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =225
                                    Top =1078
                                    Width =11775
                                    Height =3165
                                    TabIndex =6
                                    Name ="Sottomaschera Tbl_Invoices"
                                    SourceObject ="Form.SottomascheraTbl_Invoices"
                                    LinkChildFields ="Customer_ID"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_Tbl_Invoices"

                                End
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =215
                                    AccessKey =78
                                    IMESentenceMode =3
                                    Left =7350
                                    Top =6513
                                    Width =4651
                                    Height =3157
                                    TabIndex =7
                                    Name ="Testo10"
                                    UnicodeAccessKey =78

                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =223
                                            Left =7350
                                            Top =6225
                                            Width =900
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label31"
                                            Caption ="&Notes"
                                        End
                                    End
                                End
                                Begin Subform
                                    CanGrow = NotDefault
                                    OverlapFlags =247
                                    Left =225
                                    Top =6413
                                    Width =6975
                                    Height =3240
                                    TabIndex =8
                                    Name ="Sottomaschera TblNotes"
                                    SourceObject ="Form.SottomascheraTblNotes"
                                    LinkChildFields ="CustomerCode"
                                    LinkMasterFields ="Customer_code"
                                    EventProcPrefix ="Sottomaschera_TblNotes"

                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =12195
                                    Top =7597
                                    Width =2366
                                    Height =360
                                    TabIndex =9
                                    Name ="Testo14"
                                    ControlSource ="NextAppointment"
                                    AfterUpdate ="[Event Procedure]"

                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =12195
                                    Top =7200
                                    Width =2374
                                    Height =299
                                    Name ="Etichetta20"
                                    Caption ="Next appointment"
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =77
                                    Left =12210
                                    Top =5085
                                    Height =737
                                    TabIndex =10
                                    Name ="Comando13"
                                    Caption ="Save co&mment"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =109

                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =12132
                                    Top =396
                                    Width =2499
                                    Height =1747
                                    TabIndex =11
                                    Name ="Text56"
                                    ControlSource ="Note"

                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =12359
                                    Top =5896
                                    Width =1691
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =12
                                    Name ="Testo58"
                                    ControlSource ="Index"
                                    Format ="Standard"

                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =1
                                    OldBorderStyle =1
                                    OverlapFlags =215
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =14173
                                    Top =5896
                                    Width =686
                                    Height =340
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =13
                                    Name ="Testo60"
                                    ControlSource ="Update_date"
                                    Format ="General Date"

                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =4365
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta63"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =7545
                                    Top =4365
                                    Width =2025
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta64"
                                    Caption ="Total overdue (EUR)"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =4695
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta65"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =6540
                                    Top =4695
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta66"
                                    Caption ="Overdue 0--> 30 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =5040
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta67"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =6540
                                    Top =5040
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta68"
                                    Caption ="Overdue 31--> 60 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =5370
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta69"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =6555
                                    Top =5370
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta70"
                                    Caption ="Overdue 61--> 90 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =5685
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta71"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =6555
                                    Top =5685
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta72"
                                    Caption ="Overdue 91--> 180 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =9630
                                    Top =6000
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta73"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =6540
                                    Top =6000
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta74"
                                    Caption ="Overdue OVER 180 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =4380
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta75"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =1230
                                    Top =4380
                                    Width =2025
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta76"
                                    Caption ="Total overdue (USD)"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =4710
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta77"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =225
                                    Top =4710
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta78"
                                    Caption ="Overdue 0--> 30 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =5055
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta79"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =225
                                    Top =5055
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta80"
                                    Caption ="Overdue 31--> 60 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =5385
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta81"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =240
                                    Top =5385
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta82"
                                    Caption ="Overdue 61--> 90 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =5700
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta83"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =240
                                    Top =5700
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta84"
                                    Caption ="Overdue 91--> 180 DAYS"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    Left =3315
                                    Top =6015
                                    Width =2250
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta85"
                                    Caption ="&Notes"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =225
                                    Top =6015
                                    Width =3015
                                    Height =255
                                    FontSize =10
                                    FontWeight =700
                                    BackColor =-2147483633
                                    Name ="Etichetta86"
                                    Caption ="Overdue OVER 180 DAYS"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =83
                            Left =122
                            Top =380
                            Width =14753
                            Height =9306
                            Name ="&Search"
                            EventProcPrefix ="Ctl_Search"
                            UnicodeAccessKey =83
                            Begin
                                Begin CommandButton
                                    OverlapFlags =247
                                    AccessKey =83
                                    Left =4707
                                    Top =903
                                    Height =737
                                    Name ="Comando41"
                                    Caption ="&Search !"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =83

                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =285
                                    Top =870
                                    Width =3993
                                    Height =340
                                    TabIndex =1
                                    Name ="Testo42"

                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =285
                                    Top =525
                                    Width =3240
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Etichetta52"
                                    Caption ="Name o part of it"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            AccessKey =69
                            Left =122
                            Top =380
                            Width =14753
                            Height =9306
                            Name ="&Email"
                            EventProcPrefix ="Ctl_Email"
                            UnicodeAccessKey =69
                            Begin
                                Begin TextBox
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =215
                                    Top =1116
                                    Width =5390
                                    Height =345
                                    Name ="Testo47"
                                    ControlSource ="Email"

                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =5865
                                    Top =1110
                                    Width =7400
                                    Height =645
                                    TabIndex =1
                                    Name ="Text45"
                                    ControlSource ="ccEmail"

                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =215
                                    Top =1995
                                    Width =13046
                                    Height =3285
                                    TabIndex =2
                                    Name ="Text46"
                                    ControlSource ="TextEmail"

                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =215
                                    Top =793
                                    Width =3240
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Etichetta53"
                                    Caption ="Main receiver(s)"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =5839
                                    Top =793
                                    Width =3240
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Etichetta54"
                                    Caption ="Receiver(s) in cc"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =215
                                    Top =1650
                                    Width =3240
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Etichetta55"
                                    Caption ="Text"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MskScheduler2.cls"
