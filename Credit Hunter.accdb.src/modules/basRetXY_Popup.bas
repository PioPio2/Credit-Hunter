Attribute VB_Name = "basRetXY_Popup"
Option Compare Database
Option Explicit

'------------------------------------------------------------------------------------
'  API per la conversione PIXEL-TWIPS
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const DIRECTION_VERTICAL = 1
Public Const DIRECTION_HORIZONTAL = 0
'------------------------------------------------------------------------------------

'THIS CODE IS ORIGINALY WRITED BY STEPHEN LEBANS

'Alessandro Baraldi
'ik2zok@libero.it
'I modifie this Module to extract only myControl X/Y coordinate
'i CLEAN the CODE to better understend my utility
'The ORIGINAL one is much more complicated, but give much more
'options and futures.



'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97 and 2K VBA
'
'Copyright: Stephen Lebans - Lebans Holdings 1999 Ltd.  www.lebans.com
'           You may use this code in your own private or commercial applications
'           without cost. Simply leave this copyright notice in the source code.
'           You may not sell htis code by itself or as part of a collection.
'
'
'Name:      PositionFormRelativeToControl
'
'Author:    Stephen Lebans
'
' Enjoy
' Stephen Lebans

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


' GetWindow() Constants
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5

' Twips per inch
Private Const TwipsPerInch = 1440&

'  Device Parameters for GetDeviceCaps()
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Private Const BITSPIXEL = 12         '  Number of bits per pixel

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6

Public Type POINTAPI
  x As Long
  Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Horizontal and Vertical Screen resolution
Private m_Screendpi As POINTAPI

'  Questa Funzione prende spunto da una Routines di LEBANS
'  modificata appositamente per la gestione dei MENU'
Public Function Menu_XY(ctl As Access.control) As POINTAPI
    ' Window handle to our Form's Detail Section
    Dim m_hWndSection As Long
    ' For positioning window
    Dim rc As RECT
    Dim lOffSet As POINTAPI
    Dim lRet As Long

    ' Since we are turning off screen redraw ignore all errors
    On Error Resume Next

    ' Get the Window handle for the form Section containing this control
    m_hWndSection = fFindSectionhWnd(ctl)
    ' Calculate the LEFT offset for this control from the edge of the Section
    ' First calc our screen resolution
    GetScreenDPI
    ' Get window rectangle of the Section
    lRet = GetWindowRect(m_hWndSection, rc)
    ' Add to Windows Section Rectangle my ctl Deplacement
    lOffSet.x = (ctl.Left / (TwipsPerInch / m_Screendpi.x)) + rc.Left& - 1
    lOffSet.Y = ((ctl.Top + ctl.Height) / (TwipsPerInch / m_Screendpi.Y)) + rc.Top - 2
    Menu_XY = lOffSet
End Function

Private Sub GetScreenDPI()
    Dim lngDC As Long
    lngDC = GetDC(0)
    m_Screendpi.x = GetDeviceCaps(lngDC, LOGPIXELSX)
    m_Screendpi.Y = GetDeviceCaps(lngDC, LOGPIXELSY)
    lngDC = ReleaseDC(0, lngDC)
End Sub

Private Function fFindSectionhWnd(ctl As Access.control) As Long
    ' Get ListBox's hWnd
    Dim hWnd_LSB As Long
    Dim hWnd_Temp As Long

    ' Window RECT vars
    Dim rc As RECT
    Dim pt As POINTAPI

    ' Loop Counters
    Dim SectionCounter As Long
    Dim ctr As Long

    ' Which Section contains the Control?
    Select Case ctl.Section

        Case acDetail   '0
            SectionCounter = 2

        Case acHeader   '1
            SectionCounter = 1

        Case acFooter   '2
            SectionCounter = 3

        Case Else
        '  ****   NEED ERROR HANDLING! ****

    End Select

    ' Setup SectionCounter
    ' Form Header, Detail and then Footer
    ctr = 1

    ' Let's get first Child Window of the FORM
    hWnd_LSB = GetWindow(ctl.Parent.hwnd, GW_CHILD)


    ' Let's walk through every sibling window of the Form
    Do
        If fGetClassName(hWnd_LSB) = "OFormSub" Then
        ' First OFormSub is the Form's Header. We want the next next one
        ' which is the detail section
            If ctr = SectionCounter Then
                fFindSectionhWnd = hWnd_LSB
                Exit Function
            End If

             ' Increment our Section Counter
            ctr = ctr + 1

        End If

    ' Let's get the NEXT SIBLING Window
    hWnd_LSB = GetWindow(hWnd_LSB, GW_HWNDNEXT)

    ' Let's Start the process from the Top again
    ' Really just an error check
    Loop While hWnd_LSB <> 0

    ' SORRY - NO ListBox hWnd is available
    fFindSectionhWnd = 0
End Function

' From Dev Ashish's Site
' The Access Web
' http://www.mvps.org/access/

'******* Code Start *********
Private Function fGetClassName(hwnd As Long)
Dim strBuffer As String
Dim lngLen As Long
Const MAX_LEN = 255
    strBuffer = Space$(MAX_LEN)
    lngLen = GetClassName(hwnd, strBuffer, MAX_LEN)
    If lngLen > 0 Then fGetClassName = Left$(strBuffer, lngLen)
End Function
'******* Code End *********

'------------------------------------------------------------------------------------

Function fTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips
    On Error GoTo E_HANDLE
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = ReleaseDC(0, lngDeviceHandle)
    fTwipsToPixels = lngTwips / 1440 * lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_HANDLE:
    MsgBox err.Description, vbOKOnly + vbCritical, "Error: " & err.number
    Resume fExit
End Function

Function fPixelsToTwips(lngPixels As Long, lngDirection As Long) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:

'       the number of twips corresponding to the given pixels
    On Error GoTo E_HANDLE
    Dim lngDeviceHandle As Long
    Dim lngPixelsPerInch As Long
    lngDeviceHandle = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    Else
    lngPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    End If
    lngDeviceHandle = ReleaseDC(0, lngDeviceHandle)
    fPixelsToTwips = lngPixels * 1440 / lngPixelsPerInch
fExit:
    On Error Resume Next
    Exit Function
E_HANDLE:
    MsgBox err.Description, vbOKOnly + vbCritical, "Error: " & err.number
    Resume fExit
End Function
