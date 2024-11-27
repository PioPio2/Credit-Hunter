Attribute VB_Name = "basOGL"

Option Explicit
Option Compare Database
'-------------------------------------------------
'    Picture functions using GDIPlus-API (GDIP)   |
'-------------------------------------------------
'    *  Office 2007 version only!!!  *            |
'-------------------------------------------------
'   (c) mossSOFT / Sascha Trowitzsch rev. 04/2009 |
'-------------------------------------------------

'- Reference to library "OLE Automation" (stdole) needed!

'- Code does only work under Office 2007! (see *Remark below)


Public Const GUID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"    'IPicture

'User-defined types: ----------------------------------------------------------------------

Public Enum PicFileType
    pictypeBMP = 1
    pictypeGIF = 2
    pictypePNG = 3
    pictypeJPG = 4
End Enum

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type TSize
    x As Double
    Y As Double
End Type

Public Type RECT
    Bottom As Long
    Left As Long
    Right As Long
    Top As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GDIPStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    UUID As GUID
    NumberOfValues As Long
    Type As Long
    value As Long
End Type

Private Type EncoderParameters
    Count As Long
    parameter As EncoderParameter
End Type

'API-Declarations: ----------------------------------------------------------------------------

'Convert a windows bitmap to OLE-Picture :
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As Object) As Long
'Retrieve GUID-Type from string :
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pCLSID As GUID) As Long

'Memory functions:
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByRef Source As Byte, ByVal Length As Long)

'Modules API:
Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'Timer API:
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


'OLE-Stream functions :
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32.dll" (ByVal pstm As Any, ByRef phglobal As Long) As Long

'GDIPlus-API Declarations:

'*Remark:
'This uses the special gdi+ version of Office 2007! (program files\common files\microsoft shared\office12\ogl.dll)
'Benefit: No need to load a separate dll because ogl.dll is *always* loaded by Office 2007.
'ogl.dll is identical to gdiplus.dll (V1.1) used in Vista OS

'Initialization GDIP:
Private Declare PtrSafe Function GdiplusStartup Lib "ogl" (Token As Long, inputbuf As GDIPStartupInput, Optional ByVal outputbuf As Long = 0) As Long
'Tear down GDIP:
Private Declare PtrSafe Function GdiplusShutdown Lib "ogl" (ByVal Token As Long) As Long
'Load GDIP-Image from file :
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "ogl" (ByVal FileName As Long, bitmap As Long) As Long
'Create GDIP- graphical area from Windows-DeviceContext:
Private Declare PtrSafe Function GdipCreateFromHDC Lib "ogl" (ByVal hdc As Long, GpGraphics As Long) As Long
'Delete GDIP graphical area :
Private Declare PtrSafe Function GdipDeleteGraphics Lib "ogl" (ByVal Graphics As Long) As Long
'Copy GDIP-Image to graphical area:
Private Declare PtrSafe Function GdipDrawImageRect Lib "ogl" (ByVal Graphics As Long, ByVal image As Long, ByVal x As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
'Clear allocated bitmap memory from GDIP :
Private Declare PtrSafe Function GdipDisposeImage Lib "ogl" (ByVal image As Long) As Long
'Retrieve windows bitmap handle from GDIP-Image:
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "ogl" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
'Retrieve Windows-Icon-Handle from GDIP-Image:
Public Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long) As Long
'Scaling GDIP-Image size:
Private Declare PtrSafe Function GdipGetImageThumbnail Lib "ogl" (ByVal image As Long, ByVal thumbWidth As Long, ByVal thumbHeight As Long, thumbImage As Long, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
'Retrieve GDIP-Image from Windows-Bitmap-Handle:
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "ogl" (ByVal hbm As Long, ByVal hPal As Long, bitmap As Long) As Long
'Retrieve GDIP-Image from Windows-Icon-Handle:
Private Declare PtrSafe Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hicon As Long, bitmap As Long) As Long
'Retrieve width of a GDIP-Image (Pixel):
Private Declare PtrSafe Function GdipGetImageWidth Lib "ogl" (ByVal image As Long, Width As Long) As Long
'Retrieve height of a GDIP-Image (Pixel):
Private Declare PtrSafe Function GdipGetImageHeight Lib "ogl" (ByVal image As Long, Height As Long) As Long
'Save GDIP-Image to file in seletable format:
Private Declare PtrSafe Function GdipSaveImageToFile Lib "ogl" (ByVal image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
'Save GDIP-Image in OLE-Stream with seletable format:
Private Declare PtrSafe Function GdipSaveImageToStream Lib "ogl" (ByVal image As Long, ByVal stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
'Retrieve GDIP-Image from OLE-Stream-Object:
Private Declare PtrSafe Function GdipLoadImageFromStream Lib "ogl" (ByVal stream As IUnknown, image As Long) As Long


'-----------------------------------------------------------------------------------------
'Global module variable:
Private lGDIP As Long
'-----------------------------------------------------------------------------------------


'Initialize GDI+
Function InitGDIP() As Boolean
    Dim TGDP As GDIPStartupInput
    Dim hMod As Long
    
    If lGDIP = 0 Then
        If IsNull(TempVars("GDIPlusHandle")) Then   'If lGDIP is broken due to unhandled errors restore it from the Tempvars collection
            TGDP.GdiplusVersion = 1
            hMod = GetModuleHandle("ogl.dll")   'ogl.dll not yet loaded?
            If hMod = 0 Then
                hMod = LoadLibrary(Environ$("CommonProgramFiles") & "\Microsoft Shared\Office12\ogl.dll")
            End If
            GdiplusStartup lGDIP, TGDP
            TempVars("GDIPlusHandle") = lGDIP
        Else
            lGDIP = TempVars("GDIPlusHandle")
        End If
    End If
    InitGDIP = (lGDIP > 0)
    AutoShutDown
End Function

'Clear GDI+
Sub ShutDownGDIP()
    If lGDIP <> 0 Then
        GdiplusShutdown lGDIP
        lGDIP = 0
        TempVars("GDIPlusHandle") = Null
        'FreeLibrary GetModuleHandle("ogl.dll")
    End If
End Sub

'Scheduled ShutDown of GDI+ handle to avoid memory leaks
Private Sub AutoShutDown()
    'Set to 5 seconds for next shutdown
    'That's IMO appropriate for looped routines  - but configure for your own purposes
    If lGDIP <> 0 Then
        SetTimer 0&, 0&, 5000, AddressOf TimerProc
    End If
End Sub

'Callback for AutoShutDown
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Debug.Print "GDI+ AutoShutDown"
    KillTimer 0&, idEvent
    ShutDownGDIP
End Sub

'Load image file with GDIP
'It's equivalent to the method LoadPicture() in OLE-Automation library (stdole2.tlb)
'Allowed format: bmp, gif, jp(e)g, tif, png, wmf, emf, ico
Function LoadPictureGDIP(sFileName As String) As StdPicture
    Dim hBmp As Long
    Dim hPic As Long

    If Not InitGDIP Then Exit Function
    If GdipCreateBitmapFromFile(StrPtr(sFileName), hPic) = 0 Then
        GdipCreateHBITMAPFromBitmap hPic, hBmp, 0&
        If hBmp <> 0 Then
            Set LoadPictureGDIP = BitmapToPicture(hBmp)
            GdipDisposeImage hPic
        End If
    End If

End Function

'Scale picture with GDIP
'A Picture object is commited, also the return value
'Width and Height of generatrix pictures in Width, Height
'bSharpen: TRUE=Thumb is additional sharpened
Function ResampleGDIP(ByVal image As StdPicture, ByVal Width As Long, ByVal Height As Long, _
                      Optional bSharpen As Boolean = True) As StdPicture
    Dim lRes As Long
    Dim lBitmap As Long

    If Not InitGDIP Then Exit Function
    
    If image.Type = 1 Then
        lRes = GdipCreateBitmapFromHBITMAP(image.handle, 0, lBitmap)
    Else
        lRes = GdipCreateBitmapFromHICON(image.handle, lBitmap)
    End If
    If lRes = 0 Then
        Dim lThumb As Long
        Dim hBitmap As Long

        lRes = GdipGetImageThumbnail(lBitmap, Width, Height, lThumb, 0, 0)
        If lRes = 0 Then
            If image.Type = 3 Then  'Image-Type 3 is named : Icon!
                'Convert with these GDI+ method :
                lRes = GdipCreateHICONFromBitmap(lThumb, hBitmap)
                Set ResampleGDIP = BitmapToPicture(hBitmap, True)
            Else
                lRes = GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0)
                Set ResampleGDIP = BitmapToPicture(hBitmap)
            End If
            
            GdipDisposeImage lThumb
        End If
        GdipDisposeImage lBitmap
    End If

End Function

'Retrieve Width and Height of a pictures in Pixel with GDIP
'Return value as user/defined type TSize (X/Y als Long)
Function GetDimensionsGDIP(ByVal image As StdPicture) As TSize
    Dim lRes As Long
    Dim lBitmap As Long
    Dim x As Long, Y As Long

    If Not InitGDIP Then Exit Function
    If image Is Nothing Then Exit Function
    lRes = GdipCreateBitmapFromHBITMAP(image.handle, 0, lBitmap)
    If lRes = 0 Then
        GdipGetImageHeight lBitmap, Y
        GdipGetImageWidth lBitmap, x
        GetDimensionsGDIP.x = CDbl(x)
        GetDimensionsGDIP.Y = CDbl(Y)
        GdipDisposeImage lBitmap
    End If

End Function

'Save a bitmap as file (with format conversion!)
'image = StdPicture object
'sFile = complete file path
'PicType = pictypeBMP, pictypeGIF, pictypePNG oder pictypeJPG
'Quality: 0...100; (works only with pictypeJPG!)
'Returns TRUE if successful
Function SavePicGDIPlus(ByVal image As StdPicture, sFile As String, _
                        PicType As PicFileType, Optional Quality As Long = 80) As Boolean
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(image.handle, 0, lBitmap) = 0 Then
        Select Case PicType
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder
        If PicType = pictypeJPG Then
            TParams.Count = 1
            With TParams.parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .Type = 4
                .value = VarPtr(CLng(Quality))
            End With
        Else
            'Different numbers of parameter between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.Count = 1 Else TParams.Count = 0
        End If
        'Save GDIP-Image to file :
        ret = GdipSaveImageToFile(lBitmap, StrPtr(sFile), TEncoder, TParams)
        GdipDisposeImage lBitmap
        DoEvents
        'Function returns True, if generated file actually exists:
        SavePicGDIPlus = (Dir(sFile) <> "")
    End If

End Function

'This procedure is similar to the above (see Parameter), the different is,
'that nothing is stored as a file, but a conversion is executed
'using a OLE-Stream-Object to an Byte-Array .
Function ArrayFromPicture(ByVal image As Object, PicType As PicFileType, Optional Quality As Long = 80) As Byte()
    Dim lBitmap As Long
    Dim TEncoder As GUID
    Dim ret As Long
    Dim TParams As EncoderParameters
    Dim sType As String
    Dim IStm As IUnknown

    If Not InitGDIP Then Exit Function

    If GdipCreateBitmapFromHBITMAP(image.handle, 0, lBitmap) = 0 Then
        Select Case PicType    'Choose GDIP-Format-Encoders CLSID:
        Case pictypeBMP: sType = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeGIF: sType = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypePNG: sType = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
        Case pictypeJPG: sType = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
        End Select
        CLSIDFromString StrPtr(sType), TEncoder

        If PicType = pictypeJPG Then    'If JPG, set additional parameter
            'to apply the quality level
            TParams.Count = 1
            With TParams.parameter    ' Quality
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .UUID
                .NumberOfValues = 1
                .Type = 4
                .value = VarPtr(CLng(Quality))
            End With
        Else
            'Different numbers of parameter between GDI+ 1.0 and GDI+ 1.1 on GIFs!!
            If (PicType = pictypeGIF) Then TParams.Count = 1 Else TParams.Count = 0
        End If

        ret = CreateStreamOnHGlobal(0&, 1, IStm)    'Create stream
        'Save GDIP-Image to stream :
        ret = GdipSaveImageToStream(lBitmap, IStm, TEncoder, TParams)
        If ret = 0 Then
            Dim hMem As Long, LSize As Long, lpMem As Long
            Dim abData() As Byte

            ret = GetHGlobalFromStream(IStm, hMem)    'Get Memory-Handle from stream
            If ret = 0 Then
                LSize = GlobalSize(hMem)
                lpMem = GlobalLock(hMem)   'Get access to memory
                ReDim abData(LSize - 1)    'Arrays dimension
                'Commit memory stack from streams :
                CopyMemory abData(0), ByVal lpMem, LSize
                GlobalUnlock hMem   'Lock memory
                ArrayFromPicture = abData   'Result
            End If

            Set IStm = Nothing  'Clean
        End If

        GdipDisposeImage lBitmap    'Clear GDIP-Image-Memory
    End If

End Function

'Create a picture object from an Access 2007 attachment
'strTable:              Table containing picture file attachments
'strAttachmentField:    Name of the attachment column in the table
'strImage:              Name of the image to search in the attachment records
'? AttachmentToPicture("ribbonimages","imageblob","cloudy.png").Width
Public Function AttachmentToPicture(strTable As String, strAttachmentField As String, strImage As String) As StdPicture
    Dim strSQL As String
    Dim bin() As Byte
    Dim nOffset As Long
    Dim nSize As Long
    
    strSQL = "SELECT " & strTable & "." & strAttachmentField & ".FileData AS data " & _
             "FROM " & strTable & _
             " WHERE " & strTable & "." & strAttachmentField & ".FileName='" & strImage & "'"
    On Error Resume Next
    bin = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenSnapshot)(0)
    If err.number = 0 Then
        Dim bin2() As Byte
        nOffset = bin(0)    'First byte of Field2.FileData identifies offset to the file data block
        nSize = UBound(bin)
        ReDim bin2(nSize - nOffset)
        CopyMemory bin2(0), bin(nOffset), nSize - nOffset   'Copy file into new byte array starting at offset
        Set AttachmentToPicture = ArrayToPicture(bin2)
        Erase bin2
        Erase bin
    End If
End Function

'Create an OLE-Picture from Byte-Array PicBin()
Public Function ArrayToPicture(ByRef PicBin() As Byte) As StdPicture
    Dim IStm As IUnknown
    Dim lBitmap As Long
    Dim hBmp As Long
    Dim ret As Long

    If Not InitGDIP Then Exit Function

    ret = CreateStreamOnHGlobal(VarPtr(PicBin(0)), 0, IStm)  'Create stream from memory stack
    If ret = 0 Then    'OK, start GDIP :
        'Convert stream to GDIP-Image :
        ret = GdipLoadImageFromStream(IStm, lBitmap)
        If ret = 0 Then
            'Get Windows-Bitmap from GDIP-Image:
            GdipCreateHBITMAPFromBitmap lBitmap, hBmp, 0&
            If hBmp <> 0 Then
                'Convert bitmap to picture object :
                Set ArrayToPicture = BitmapToPicture(hBmp)
            End If
        End If
        'Clear memory ...
        GdipDisposeImage lBitmap
    End If

End Function

'Help function to get a OLE-Picture from Windows-Bitmap-Handle
'If bIsIcon = TRUE, an Icon-Handle is commited
Function BitmapToPicture(ByVal hBmp As Long, Optional bIsIcon As Boolean = False) As StdPicture
    Dim TPicConv As PICTDESC, UID As GUID

    With TPicConv
        If bIsIcon Then
            .cbSizeOfStruct = 16
            .PicType = 3    'PicType Icon
        Else
            .cbSizeOfStruct = Len(TPicConv)
            .PicType = 1    'PicType Bitmap
        End If
        .hImage = hBmp
    End With

    CLSIDFromString StrPtr(GUID_IPicture), UID
    OleCreatePictureIndirect TPicConv, UID, True, BitmapToPicture

End Function

