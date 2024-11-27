Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6426
    DatasheetFontHeight =10
    ItemSuffix =3
    Right =15351
    Bottom =9659
    RecSrcDt = Begin
        0x2401b8aed48ce340
    End
    RecordSource ="TblGeneral"
    Caption ="CLs receiver E-mail addresses"
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
        Begin Section
            Height =2437
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =285
                    Top =1250
                    Width =5777
                    Height =1069
                    Name ="ToBeSentCLto"
                    ControlSource ="ToBeSentCLto"

                End
                Begin Label
                    OverlapFlags =85
                    Left =285
                    Top =258
                    Width =5774
                    Height =493
                    Name ="Label1"
                    Caption ="Please insert your email addresses in the field below if you want to receive aut"
                        "omatically by email the CL report every morning."
                End
                Begin Label
                    OverlapFlags =85
                    Left =285
                    Top =869
                    Width =5950
                    Height =218
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Note that different email addresses must be separated by a comma \",\""
                End
            End
        End
    End
End
CodeBehindForm
' See "MskCLsE-mailReceiverAddresses.cls"
