Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7920
    DatasheetFontHeight =10
    ItemSuffix =56
    Left =210
    Top =5280
    Right =14505
    Bottom =8400
    RecSrcDt = Begin
        0x0bf3a82c8a3ae640
    End
    RecordSource ="SELECT Tbl_credit_check_failures.ID, Tbl_credit_check_failures.[Hold Type], Tbl_"
        "credit_check_failures.[Hold Name], Tbl_credit_check_failures.[Date Hold Applied]"
        ", Tbl_credit_check_failures.[Hold Until Date], Tbl_credit_check_failures.[Hold C"
        "omments], Tbl_credit_check_failures.[Sub-Region], Tbl_credit_check_failures.Coun"
        "try, Tbl_credit_check_failures.[Customer Name], Tbl_credit_check_failures.[Custo"
        "mer Number], Tbl_credit_check_failures.[Account Specialist], Tbl_credit_check_fa"
        "ilures.[Logitech Item Number], Tbl_credit_check_failures.[List Price], Tbl_credi"
        "t_check_failures.[Requested Quantity], Tbl_credit_check_failures.[Currency Code]"
        ", Tbl_credit_check_failures.Amount, Tbl_credit_check_failures.[Order Number], Tb"
        "l_credit_check_failures.[Order Line Number], Tbl_credit_check_failures.[Order Da"
        "te], Tbl_credit_check_failures.[Requested Date], Tbl_credit_check_failures.[Sche"
        "dule Date], Tbl_credit_check_failures.[Active Hold], Tbl_credit_check_failures.["
        "Open Line], Tbl_credit_check_failures.[Line Status], Tbl_credit_check_failures.["
        "Hold Criteria], Tbl_credit_check_failures.[Tax Code], Tbl_credit_check_failures."
        "Released FROM Tbl_credit_check_failures ORDER BY Tbl_credit_check_failures.[Orde"
        "r Number], Tbl_credit_check_failures.[Order Line Number]; "
    Caption ="Sottomaschera Tbl_credit_check_failures"
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
            Height =4269
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =456
                    Width =1821
                    Height =450
                    ColumnWidth =1950
                    TabIndex =1
                    Name ="HoldType"
                    ControlSource ="Hold Type"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =456
                            Width =684
                            Height =255
                            Name ="Hold Type_Etichetta"
                            Caption ="Hold Type"
                            EventProcPrefix ="Hold_Type_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =969
                    Width =1821
                    Height =450
                    ColumnWidth =1950
                    TabIndex =2
                    Name ="HoldName"
                    ControlSource ="Hold Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =969
                            Width =684
                            Height =255
                            Name ="Hold Name_Etichetta"
                            Caption ="Hold Name"
                            EventProcPrefix ="Hold_Name_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =1482
                    Width =1035
                    Height =255
                    ColumnWidth =1740
                    TabIndex =3
                    Name ="DateHoldApplied"
                    ControlSource ="Date Hold Applied"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1482
                            Width =684
                            Height =255
                            Name ="Date Hold Applied_Etichetta"
                            Caption ="Date Hold Applied"
                            EventProcPrefix ="Date_Hold_Applied_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =1824
                    Width =1035
                    Height =255
                    ColumnWidth =1470
                    TabIndex =4
                    Name ="HoldUntilDate"
                    ControlSource ="Hold Until Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1824
                            Width =684
                            Height =255
                            Name ="Hold Until Date_Etichetta"
                            Caption ="Hold Until Date"
                            EventProcPrefix ="Hold_Until_Date_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =2166
                    Width =1821
                    Height =450
                    ColumnWidth =7335
                    TabIndex =5
                    Name ="HoldComments"
                    ControlSource ="Hold Comments"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2166
                            Width =684
                            Height =255
                            Name ="Hold Comments_Etichetta"
                            Caption ="Hold Comments"
                            EventProcPrefix ="Hold_Comments_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =2679
                    Width =1821
                    Height =450
                    ColumnWidth =3000
                    TabIndex =6
                    Name ="SubRegion"
                    ControlSource ="Sub-Region"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2679
                            Width =684
                            Height =255
                            Name ="Sub-Region_Etichetta"
                            Caption ="Sub-Region"
                            EventProcPrefix ="Sub_Region_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =3192
                    Width =1821
                    Height =450
                    ColumnWidth =3000
                    TabIndex =7
                    Name ="Country"
                    ControlSource ="Country"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =3192
                            Width =684
                            Height =255
                            Name ="Country_Etichetta"
                            Caption ="Country"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =801
                    Top =3705
                    Width =1821
                    Height =450
                    ColumnWidth =3000
                    TabIndex =8
                    Name ="CustomerName"
                    ControlSource ="Customer Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =3705
                            Width =684
                            Height =255
                            Name ="Customer Name_Etichetta"
                            Caption ="Customer Name"
                            EventProcPrefix ="Customer_Name_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =114
                    Width =1821
                    Height =255
                    ColumnWidth =1755
                    TabIndex =9
                    Name ="CustomerNumber"
                    ControlSource ="Customer Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =114
                            Width =684
                            Height =255
                            Name ="Customer Number_Etichetta"
                            Caption ="Customer Number"
                            EventProcPrefix ="Customer_Number_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =456
                    Width =1821
                    Height =450
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="AccountSpecialist"
                    ControlSource ="Account Specialist"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =456
                            Width =684
                            Height =255
                            Name ="Account Specialist_Etichetta"
                            Caption ="Account Specialist"
                            EventProcPrefix ="Account_Specialist_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =969
                    Width =1821
                    Height =450
                    ColumnWidth =2085
                    TabIndex =11
                    Name ="LogitechItemNumber"
                    ControlSource ="Logitech Item Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =969
                            Width =684
                            Height =255
                            Name ="Logitech Item Number_Etichetta"
                            Caption ="Logitech Item Number"
                            EventProcPrefix ="Logitech_Item_Number_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =1482
                    Width =1821
                    Height =255
                    ColumnWidth =1620
                    TabIndex =12
                    Name ="ListPrice"
                    ControlSource ="List Price"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =1482
                            Width =684
                            Height =255
                            Name ="List Price_Etichetta"
                            Caption ="List Price"
                            EventProcPrefix ="List_Price_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =1824
                    Width =1821
                    Height =255
                    ColumnWidth =1905
                    TabIndex =13
                    Name ="RequestedQuantity"
                    ControlSource ="Requested Quantity"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =1824
                            Width =684
                            Height =255
                            Name ="Requested Quantity_Etichetta"
                            Caption ="Requested Quantity"
                            EventProcPrefix ="Requested_Quantity_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =2166
                    Width =1821
                    Height =450
                    ColumnWidth =1470
                    TabIndex =14
                    Name ="CurrencyCode"
                    ControlSource ="Currency Code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =2166
                            Width =684
                            Height =255
                            Name ="Currency Code_Etichetta"
                            Caption ="Currency Code"
                            EventProcPrefix ="Currency_Code_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =2679
                    Width =1821
                    Height =255
                    ColumnWidth =2310
                    TabIndex =15
                    Name ="Amount"
                    ControlSource ="Amount"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =2679
                            Width =684
                            Height =255
                            Name ="Amount_Etichetta"
                            Caption ="Amount"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =3021
                    Width =1821
                    Height =255
                    ColumnWidth =2310
                    TabIndex =16
                    Name ="OrderNumber"
                    ControlSource ="Order Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =3021
                            Width =684
                            Height =255
                            Name ="Order Number_Etichetta"
                            Caption ="Order Number"
                            EventProcPrefix ="Order_Number_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3423
                    Top =3363
                    Width =1821
                    Height =450
                    ColumnWidth =1815
                    TabIndex =17
                    Name ="OrderLineNumber"
                    ControlSource ="Order Line Number"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2679
                            Top =3363
                            Width =684
                            Height =255
                            Name ="Order Line Number_Etichetta"
                            Caption ="Order Line Number"
                            EventProcPrefix ="Order_Line_Number_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =114
                    Width =1035
                    Height =255
                    ColumnWidth =1560
                    TabIndex =18
                    Name ="OrderDate"
                    ControlSource ="Order Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =114
                            Width =684
                            Height =255
                            Name ="Order Date_Etichetta"
                            Caption ="Order Date"
                            EventProcPrefix ="Order_Date_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =456
                    Width =1035
                    Height =255
                    ColumnWidth =1695
                    TabIndex =19
                    Name ="RequestedDate"
                    ControlSource ="Requested Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =456
                            Width =684
                            Height =255
                            Name ="Requested Date_Etichetta"
                            Caption ="Requested Date"
                            EventProcPrefix ="Requested_Date_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =798
                    Width =1035
                    Height =255
                    ColumnWidth =1695
                    TabIndex =20
                    Name ="ScheduleDate"
                    ControlSource ="Schedule Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =798
                            Width =684
                            Height =255
                            Name ="Schedule Date_Etichetta"
                            Caption ="Schedule Date"
                            EventProcPrefix ="Schedule_Date_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =1140
                    Width =1818
                    Height =450
                    ColumnWidth =1395
                    TabIndex =21
                    Name ="ActiveHold"
                    ControlSource ="Active Hold"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =1140
                            Width =684
                            Height =255
                            Name ="Active Hold_Etichetta"
                            Caption ="Active Hold"
                            EventProcPrefix ="Active_Hold_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =1653
                    Width =1818
                    Height =450
                    ColumnWidth =1305
                    TabIndex =22
                    Name ="OpenLine"
                    ControlSource ="Open Line"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =1653
                            Width =684
                            Height =255
                            Name ="Open Line_Etichetta"
                            Caption ="Open Line"
                            EventProcPrefix ="Open_Line_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =2166
                    Width =1818
                    Height =450
                    ColumnWidth =1725
                    TabIndex =23
                    Name ="LineStatus"
                    ControlSource ="Line Status"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =2166
                            Width =684
                            Height =255
                            Name ="Line Status_Etichetta"
                            Caption ="Line Status"
                            EventProcPrefix ="Line_Status_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =2679
                    Width =1818
                    Height =450
                    ColumnWidth =1245
                    TabIndex =24
                    Name ="HoldCriteria"
                    ControlSource ="Hold Criteria"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =2679
                            Width =684
                            Height =255
                            Name ="Hold Criteria_Etichetta"
                            Caption ="Hold Criteria"
                            EventProcPrefix ="Hold_Criteria_Etichetta"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6045
                    Top =3192
                    Width =1818
                    Height =450
                    ColumnWidth =1005
                    TabIndex =25
                    Name ="TaxCode"
                    ControlSource ="Tax Code"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5301
                            Top =3192
                            Width =684
                            Height =255
                            Name ="Tax Code_Etichetta"
                            Caption ="Tax Code"
                            EventProcPrefix ="Tax_Code_Etichetta"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1757
                    Top =170
                    ColumnWidth =1875
                    Name ="Check54"
                    ControlSource ="Released"

                    LayoutCachedLeft =1757
                    LayoutCachedTop =170
                    LayoutCachedWidth =2017
                    LayoutCachedHeight =410
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1140
                            Height =240
                            Name ="Label55"
                            Caption ="To be released"
                            LayoutCachedLeft =120
                            LayoutCachedTop =120
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =360
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
CodeBehindForm
' See "SottomascheraTbl_credit_check_failures.cls"
