Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5504
    RowHeight =2970
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =690
    Top =9585
    Right =10905
    Bottom =12150
    RecSrcDt = Begin
        0x7884ac2c8a3ae640
    End
    RecordSource ="SELECT TblNotes.*, TblNotes.DateTimeOfTheNote, TblNotes.ID FROM TblNotes ORDER B"
        "Y TblNotes.DateTimeOfTheNote DESC , TblNotes.ID DESC; "
    Caption ="Sottomaschera TblNotes"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Moveable =0
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
        Begin CustomControl
            SpecialEffect =2
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
            Height =2729
            BackColor =-2147483633
            Name ="Corpo"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    ScrollBarAlign =1
                    IMESentenceMode =3
                    Left =56
                    Width =5448
                    Height =2729
                    ColumnWidth =18900
                    Name ="Note"
                    ControlSource ="Note"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =56
                    LayoutCachedWidth =5504
                    LayoutCachedHeight =2729
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
' See "SottomascheraTblNotes.cls"
