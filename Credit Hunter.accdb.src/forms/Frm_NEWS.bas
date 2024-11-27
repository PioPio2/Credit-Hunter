Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9018
    DatasheetFontHeight =10
    ItemSuffix =56
    Left =9195
    Top =8325
    Right =18195
    Bottom =11850
    RecSrcDt = Begin
        0x7930285eb01ee340
    End
    RecordSource ="Tbl_NEWS"
    Caption ="Tbl_NEWS"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnTimer ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =480
            Name ="IntestazioneMaschera"
            OnMouseMove ="[Event Procedure]"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2224
                    Top =45
                    Width =2313
                    FontWeight =700
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="lbl_tit"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2224
                    LayoutCachedTop =45
                    LayoutCachedWidth =4537
                    LayoutCachedHeight =285
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2835
                    Left =5569
                    Top =48
                    Width =2415
                    TabIndex =1
                    BackColor =16316664
                    BorderColor =9868950
                    ForeColor =12615680
                    Name ="lbl_opz"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Tbl_RSS.ID, Tbl_RSS.RSS_TITLE, Tbl_RSS.RSS_ADDRESS FROM Tbl_RSS; "
                    ColumnWidths ="0;2835;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =5569
                    LayoutCachedTop =48
                    LayoutCachedWidth =7984
                    LayoutCachedHeight =288
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4657
                    Top =48
                    Width =804
                    Height =240
                    ForeColor =26367
                    Name ="Etichetta29"
                    Caption ="Browse"
                    FontName ="Tahoma"
                    LayoutCachedLeft =4657
                    LayoutCachedTop =48
                    LayoutCachedWidth =5461
                    LayoutCachedHeight =288
                End
                Begin Line
                    OverlapFlags =85
                    Left =31
                    Top =330
                    Width =8888
                    BorderColor =12632256
                    Name ="Linea30"
                    LayoutCachedLeft =31
                    LayoutCachedTop =330
                    LayoutCachedWidth =8919
                    LayoutCachedHeight =330
                End
                Begin Line
                    OverlapFlags =85
                    Left =4621
                    Top =56
                    Width =0
                    Height =227
                    BorderColor =12632256
                    Name ="Linea31"
                    LayoutCachedLeft =4621
                    LayoutCachedTop =56
                    LayoutCachedWidth =4621
                    LayoutCachedHeight =283
                End
                Begin Image
                    BackStyle =1
                    Left =8319
                    Top =90
                    Width =180
                    Height =180
                    Name ="lbl_giu_over"
                    PictureData = Begin
                        0x280000000c0000000c000000010008000000000090000000c40e0000c40e0000 ,
                        0x0001000000010000dddddd003232ff00ffffff00cbcbcb00eeeeee00ffffff00 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000050303030303030303030305030202020202020202020203 ,
                        0x0302040404040404040400030302040404010104040400030302040401010101 ,
                        0x0404000303020401010101010104000303020404040101040404000303020404 ,
                        0x0401010404040003030204040401010404040003030204040404040404040003 ,
                        0x030200000000000000000003050303030303030303030305
                    End
                    OnMouseMove ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x00030100dddddd0000000000
                    End

                    LayoutCachedLeft =8319
                    LayoutCachedTop =90
                    LayoutCachedWidth =8499
                    LayoutCachedHeight =270
                    TabIndex =5
                End
                Begin Image
                    BackStyle =1
                    Left =8100
                    Top =94
                    Width =180
                    Height =180
                    Name ="lbl_su_over"
                    PictureData = Begin
                        0x280000000c0000000c000000010008000000000090000000c40e0000c40e0000 ,
                        0x0001000000010000dddddd003232ff00ffffff00cbcbcb00eeeeee00ffffff00 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000050303030303030303030305030202020202020202020203 ,
                        0x0302040404040404040400030302040404010104040400030302040404010104 ,
                        0x0404000303020404040101040404000303020401010101010104000303020404 ,
                        0x0101010104040003030204040401010404040003030204040404040404040003 ,
                        0x030200000000000000000003050303030303030303030305
                    End
                    OnMouseMove ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x00030100dddddd0000000000
                    End

                    LayoutCachedLeft =8100
                    LayoutCachedTop =94
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =274
                    TabIndex =4
                End
                Begin Image
                    Visible = NotDefault
                    BackStyle =1
                    Left =8325
                    Top =90
                    Width =180
                    Height =180
                    Name ="lbl_giu"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000c0000000c000000010008000000000090000000c40e0000c40e0000 ,
                        0x0001000000010000dddddd0073282700ffffff00cbcbcb00eeeeee00ffffff00 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000050303030303030303030305030202020202020202020203 ,
                        0x0302040404040404040400030302040404010104040400030302040401010101 ,
                        0x0404000303020401010101010104000303020404040101040404000303020404 ,
                        0x0401010404040003030204040401010404040003030204040404040404040003 ,
                        0x030200000000000000000003050303030303030303030305
                    End
                    ObjectPalette = Begin
                        0x00030100dddddd0000000000
                    End

                    LayoutCachedLeft =8325
                    LayoutCachedTop =90
                    LayoutCachedWidth =8505
                    LayoutCachedHeight =270
                    TabIndex =3
                End
                Begin Image
                    Visible = NotDefault
                    BackStyle =1
                    Left =8100
                    Top =90
                    Width =180
                    Height =180
                    Name ="lbl_su"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000c0000000c000000010008000000000090000000c40e0000c40e0000 ,
                        0x0001000000010000dddddd0073282700ffffff00cbcbcb00eeeeee00ffffff00 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000050303030303030303030305030202020202020202020203 ,
                        0x0302040404040404040400030302040404010104040400030302040404010104 ,
                        0x0404000303020404040101040404000303020401010101010104000303020404 ,
                        0x0101010104040003030204040401010404040003030204040404040404040003 ,
                        0x030200000000000000000003050303030303030303030305
                    End
                    ObjectPalette = Begin
                        0x00030100dddddd0000000000
                    End

                    LayoutCachedLeft =8100
                    LayoutCachedTop =90
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =270
                    TabIndex =2
                End
                Begin Line
                    OverlapFlags =85
                    Left =31
                    Top =465
                    Width =8888
                    BorderColor =12632256
                    Name ="Linea41"
                    LayoutCachedLeft =31
                    LayoutCachedTop =465
                    LayoutCachedWidth =8919
                    LayoutCachedHeight =465
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =93
                    Left =31
                    Top =360
                    Width =8888
                    Height =80
                    BorderColor =12632256
                    Name ="lbl_sfo"
                    LayoutCachedLeft =31
                    LayoutCachedTop =360
                    LayoutCachedWidth =8919
                    LayoutCachedHeight =440
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =87
                    Left =30
                    Top =360
                    Width =0
                    Height =80
                    BackColor =15260362
                    BorderColor =12632256
                    Name ="lbl_progress"
                    LayoutCachedLeft =30
                    LayoutCachedTop =360
                    LayoutCachedWidth =30
                    LayoutCachedHeight =440
                End
                Begin Image
                    BackStyle =1
                    Left =1950
                    Top =56
                    Width =210
                    Height =210
                    Name ="Img_webOLD"
                    PictureData = Begin
                        0x280000000e0000000e0000000100080000000000e0000000c40e0000c40e0000 ,
                        0x00010000000100002573e3000b4fd1002367dd0076aae8003a9cfa00238af300 ,
                        0x1e85ec008fb8ec005999e900a6ccf1002880ed00b9d6f5004996e900358be600 ,
                        0x93c6f8004d8ce3001c7ee000a4d0fa00c4def700d4e6f80071a0e3004684dc00 ,
                        0x2959bc003193f700215cd90052a3f200447dda001152ce002658bf00004ddc00 ,
                        0xffffff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000001f1c1d1d1d1d1d1d1d1d1d1d1c1f00001c06050505060a00 ,
                        0x00021818181c00000110091e1105060c08001a0f181d000001101e1e1e051007 ,
                        0x1e0f141e031d00000110091e110510091e0d141e081d0000010a060605060d1e ,
                        0x0b10031e0c1d0000010a0d05060c121e1910071e0a1d000001000f07091e1e0e ,
                        0x06081e0e061d00000100151e1e1203170c131e04051d00000100000f0f0d0d03 ,
                        0x131e1104041d00000102151403070b1e1e110404041d000001021a1e1e1e1e0b ,
                        0x0c170404041d000016021803140f000a0a051717171c00001f161b1b1b1b1b1b ,
                        0x1b1b1b1b1c1f0000
                    End
                    ObjectPalette = Begin
                        0x00030100e373250000000000
                    End

                    LayoutCachedLeft =1950
                    LayoutCachedTop =56
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =266
                    TabIndex =6
                End
                Begin Label
                    OverlapFlags =85
                    Left =150
                    Top =45
                    Width =1551
                    Height =276
                    FontWeight =600
                    BackColor =16777215
                    ForeColor =0
                    Name ="Label31"
                    Caption ="News of the day"
                    FontName ="Tahoma"
                    LayoutCachedLeft =150
                    LayoutCachedTop =45
                    LayoutCachedWidth =1701
                    LayoutCachedHeight =321
                End
            End
        End
        Begin Section
            Height =270
            BackColor =16249326
            Name ="Corpo"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =330
                    Top =15
                    Width =8571
                    Height =225
                    ColumnWidth =3000
                    BackColor =-2147483633
                    BorderColor =8421504
                    ForeColor =6697728
                    Name ="TITOLO"
                    ControlSource ="TITOLO"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =330
                    LayoutCachedTop =15
                    LayoutCachedWidth =8901
                    LayoutCachedHeight =240
                End
                Begin Image
                    BackStyle =1
                    Left =60
                    Top =30
                    Width =225
                    Height =225
                    Name ="Img_web"
                    PictureData = Begin
                        0x280000000d0000000d0000000100080000000000d0000000c40e0000c40e0000 ,
                        0x000100000001000000000000ffffff00fffeff00fcf3f000eedbd300f2e9e500 ,
                        0xefe6e200e8d2c600ead1c100e9dcd400eaddd500f6f1ee00d1a38400cfa38400 ,
                        0xd5a98a00d9b49a00d5b19900debca400e4cab900e2cbbb00f5ebe400f3e9e200 ,
                        0xc48e6500c38d6400c58f6600c38c6500ca967100cb9a7400cc9c7800cfa28100 ,
                        0xcea18000d0a58400d2a78600d8b39700d7b29600ddbaa000d9b9a200d8b9a200 ,
                        0xe4c4ad00dfc0a900e8d3c400e7d2c300ecdbce00e7d7cb00e9d9cd00c58e6100 ,
                        0xc58e6300c68f6400c38e6300c48f6400c28f6400c18e6300c18d6400c28e6500 ,
                        0xc3916700c9976d00c8966c00c4936b00c6956d00c8976f00c79a7400ca9d7700 ,
                        0xcda07b00d1a78400cba28100d0a78700d4ad8d00d1ab8d00d7b19300d4b09200 ,
                        0xd9b59700d7b59800d6b49700d7b69c00d7b89f00dbbda400dfc4af00dec3ae00 ,
                        0xdfc6b200e3cab600e0c8b400dec6b200e1cdbc00e7d3c200e4d1c200e7d4c500 ,
                        0xebdbce00ede1d700fdf7f200f5efea00c28e6000c38f6100c4906200c18f6100 ,
                        0xc08f6100bf8e6200c9996f00d0aa8800d5b08e00d6b69900e1c5ad00e2c8b000 ,
                        0xe7cdb500dfc7b100e4ccb600ebddd100f4ece500f2eae300c5986d00d2af8e00 ,
                        0xd9bea400dec3a900e2cdb800f5ece300f4ebe200fffefd00e2d2c100e8dbcd00 ,
                        0xf5ece200f3ece300f2ebe200fefbf700fffdfa00f3ede200fdfaf500fffffe00 ,
                        0xfefffd00fcfffd00fdffff00ffffff0000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000001730b0572717771716b597a010000007f2c702421104822 ,
                        0x0f4a68557c00000006661c17355a305b5f311a4e0a0000000445185d335c2d31 ,
                        0x17305b0c070000002a1e343517020234332f303b080000007520362f3333017d ,
                        0x3233303a53000000566d3d6c2e305e017f163037740000002a490e1d1b390201 ,
                        0x5c33173a54000000564b63623f80013c3819323729000000694c11254642611f ,
                        0x3e601641280000001552646f6e23474443400d65570000007309134f50516726 ,
                        0x274d122b79000000017d586a721478767b6a037e80000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    TabIndex =1
                End
                Begin Line
                    OldBorderStyle =4
                    OverlapFlags =85
                    BorderLineStyle =3
                    Top =255
                    Width =6336
                    BorderColor =15523543
                    Name ="Linea43"
                End
            End
        End
        Begin FormFooter
            Height =306
            Name ="PièDiPaginaMaschera"
            Begin
                Begin Image
                    BackStyle =1
                    Left =30
                    Top =45
                    Width =255
                    Height =240
                    Name ="NonAssociatoOLE52"
                    PictureData = Begin
                        0x280000001100000010000000010018000000000040030000c40e0000c40e0000 ,
                        0x0000000000000000fefefefffffffcfcfcfffffffffffffafafac7c7c7a6a6a6 ,
                        0xa1a1a1a3a3a3c3c3c3f9f9f9fcfcfcfffffffafafaffffffffffff00fffffffc ,
                        0xfcfcffffffffffffadadad4040403434344141414141413131312424243c3c3c ,
                        0xafafaffffffffffffffdfdfdffffff00fbfbfbfafafafcfcfc7f7f7f2b2b2b8c ,
                        0x8c8ce6e6e6edededdcdcdcfcfcfcd8d8d86a6a6a141414737373ffffffffffff ,
                        0xfafafa00ffffffffffff929292404040d7d7d7eaeaeaeaeaeadadadaacacacf9 ,
                        0xf9f9d0d0d0e3e3e3c3c3c31d1d1d7c7c7cfffffffdfdfd00fdfdfddcdcdc4444 ,
                        0x44c9c9c9ffffffc8c8c8d4d4d4fafafaf2f2f2fefefec1c1c1c7c7c7ffffffb4 ,
                        0xb4b41d1d1db6b6b6fefefe00ffffff7f7f7f888888c1c1c1c1c1c1fffffff5f5 ,
                        0xf5e7e7e7c8c8c8eeeeeefcfcfcf9f9f9bbbbbbd9d9d96161615b5b5bffffff00 ,
                        0xfdfdfd666666d7d7d7d9d9d9d3d3d3f5f5f5f6f6f6dbdbdba4a4a4e6e6e6f4f4 ,
                        0xf4f6f6f6d0d0d0ecececbbbbbb353535e4e4e400e4e4e46b6b6be0e0e0dbdbdb ,
                        0xfbfbfbf0f0f0fafafab5b5b5898989cacacaf6f6f6eeeeeef8f8f8dadadad4d4 ,
                        0xd43d3d3dcdcdcd00dfdfdf767676c4c4c4afafaff9f9f9f5f5f5bcbcbc535353 ,
                        0x4242427d7d7de2e2e2efefefecececb4b4b4bfbfbf494949d2d2d200ebebeb77 ,
                        0x7777e9e9e9f2f2f2eaeaea8d8d8d737373c8c8c8e0e0e0b7b7b7878787bcbcbc ,
                        0xe9e9e9f2f2f2d6d6d6444444e4e4e400ffffff767676b9b9b9bdbdbda3a3a3cd ,
                        0xcdcdf5f5f5efefefdededef0f0f0e5e5e5b6b6b6848484d4d4d49f9f9f6f6f6f ,
                        0xffffff00fefefebababa777777d4d4d4d2d2d2f4f4f4f2f2f2e8e8e8e1e1e1e9 ,
                        0xe9e9edededfafafacbcbcbd4d4d4515151c0c0c0ffffff00fffffffbfbfb7878 ,
                        0x78999999ffffffb2b2b2d4d4d4dcdcdcc1c1c1eeeeeecbcbcbd5d5d5ffffff79 ,
                        0x7979828282ffffffffffff00fafafaffffffefefef717171898989cbcbcbfafa ,
                        0xfacececea3a3a3f3f3f3f6f6f6cacaca6a6a6a7b7b7bfffffffefefefcfcfc00 ,
                        0xf8f8f8fcfcfcfefefef5f5f58888886c6c6c919191aaaaaaacacaca4a4a48181 ,
                        0x815a5a5a979797fbfbfbfcfcfcfffffffdfdfd00fafafafefefef9f9f9fdfdfd ,
                        0xffffffd8d8d8a4a4a49b9b9ba7a7a79a9a9aabababe9e9e9fffffffafafafdfd ,
                        0xfdfffffffefefe00
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    TabIndex =2
                End
                Begin Label
                    OverlapFlags =85
                    Left =330
                    Top =45
                    Width =3930
                    Height =210
                    ForeColor =9868950
                    Name ="lbl_agg"
                    Caption ="News updated as of:"
                    FontName ="Tahoma"
                End
                Begin Line
                    OverlapFlags =85
                    Top =15
                    Width =8888
                    BorderColor =12632256
                    Name ="Linea51"
                    LayoutCachedTop =15
                    LayoutCachedWidth =8888
                    LayoutCachedHeight =15
                End
                Begin Image
                    BackStyle =1
                    Left =8570
                    Top =30
                    Width =270
                    Height =270
                    Name ="Img_Reconnect"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000001000000010000000010008000000000000010000c40e0000c40e0000 ,
                        0x0001000000010000c0c0c000ffff0000800000000080800080808000ff000000 ,
                        0x00ffff00ffffff0000000000ffffff0000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000090909090909090909090909090909090909080808080808 ,
                        0x0808080808080909090904070707070707070700040809090909040404040404 ,
                        0x0404040808080909090909040707070707070704080909090909090407020505 ,
                        0x0505070408090909090909040702010505050704080909090909090407020502 ,
                        0x0202070408090909090909040707070707070704080909090909090904040403 ,
                        0x0404040909090909090909090909040708090909090909090908060908090407 ,
                        0x0809080906080909080603080609030803090608030608090806080603090807 ,
                        0x0809030608060809080603080609030803090608030608090908060908090909 ,
                        0x0909080906080909
                    End
                    ObjectPalette = Begin
                        0x00030100c0c0c00000000000
                    End
                    ControlTipText ="Update news"

                    LayoutCachedLeft =8570
                    LayoutCachedTop =30
                    LayoutCachedWidth =8840
                    LayoutCachedHeight =300
                    TabIndex =3
                End
                Begin TextBox
                    Visible = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8170
                    Top =56
                    Width =279
                    Height =210
                    FontSize =7
                    ForeColor =26367
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8170
                    LayoutCachedTop =56
                    LayoutCachedWidth =8449
                    LayoutCachedHeight =266
                End
                Begin TextBox
                    Visible = NotDefault
                    EnterKeyBehavior = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7830
                    Top =56
                    Width =279
                    Height =210
                    ColumnWidth =3000
                    FontSize =7
                    TabIndex =1
                    ForeColor =26367
                    Name ="WEB"
                    ControlSource ="WEB"
                    FontName ="Tahoma"

                    LayoutCachedLeft =7830
                    LayoutCachedTop =56
                    LayoutCachedWidth =8109
                    LayoutCachedHeight =266
                End
            End
        End
    End
End
CodeBehindForm
' See "Frm_NEWS.cls"
