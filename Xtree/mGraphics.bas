Attribute VB_Name = "mGraphics"
Option Explicit
Public Type POINTAPI
    x                                                 As Long
    y                                                 As Long
End Type
Public Type RECT
    left                                      As Long
    top                                       As Long
    right                                     As Long
    bottom                                    As Long
End Type

Private Const BITSPIXEL                           As Integer = 12
Public Const PS_SOLID                             As Integer = 0
'-- Line functions:
Private Const BF_LEFT                             As Long = &H1
Private Const BF_BOTTOM                           As Long = &H8
Private Const BF_RIGHT                            As Long = &H4
Private Const BF_TOP                              As Long = &H2
Private Const BDR_RAISEDINNER                     As Long = &H4
Private Const BDR_RAISEDOUTER                     As Long = &H1
Private Const BDR_SUNKENINNER                     As Long = &H8
Private Const BDR_SUNKENOUTER                     As Long = &H2
Private Const EDGE_BUMP                           As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED                         As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED                         As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN                         As Double = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
'-- Colour functions:
Private Const OPAQUE                              As Integer = 2
Public Const TRANSPARENT                          As Integer = 1
Public Type BITMAP
    bmType                                            As Long
    bmWidth                                           As Long
    bmHeight                                          As Long
    bmWidthBytes                                      As Long
    bmPlanes                                          As Long
    bmBitsPixel                                       As Integer
    bmBits                                            As Long
End Type
Public Const DT_LEFT                              As Long = &H0
Public Const DT_RIGHT                             As Long = &H2
Public Const DT_END_ELLIPSIS                      As Long = &H8000&
'-- Scrolling and region functions:
Public Const LF_FACESIZE                          As Integer = 32
Public Type LOGFONT
    lfHeight                                          As Long
    lfWidth                                           As Long
    lfEscapement                                      As Long
    lfOrientation                                     As Long
    lfWeight                                          As Long
    lfItalic                                          As Byte
    lfUnderline                                       As Byte
    lfStrikeOut                                       As Byte
    lfCharSet                                         As Byte
    lfOutPrecision                                    As Byte
    lfClipPrecision                                   As Byte
    lfQuality                                         As Byte
    lfPitchAndFamily                                  As Byte
    lfFaceName(LF_FACESIZE)                           As Byte
End Type
Public Const DST_ICON                             As Long = &H3
Public Const DSS_DISABLED                         As Long = &H20
Public Const DSS_MONO                             As Long = &H80
#If False Then
    Private DFCS_CAPTIONCLOSE, DFCS_CAPTIONMIN, DFCS_CAPTIONMAX, DFCS_CAPTIONRESTORE, DFCS_CAPTIONHELP
#End If
Public Enum DFCMenuTypeFlags
'-- Menu types:
    DFCS_MENUARROW = &H0&
    DFCS_MENUCHECK = &H1&
    DFCS_MENUBULLET = &H2&
    DFCS_MENUARROWRIGHT = &H4&
End Enum
#If False Then
    Private DFCS_MENUARROW, DFCS_MENUCHECK, DFCS_MENUBULLET, DFCS_MENUARROWRIGHT
#End If
Public Enum DFCScrollTypeFlags
'-- Scroll types:
    DFCS_SCROLLUP = &H0&
    DFCS_SCROLLDOWN = &H1&
    DFCS_SCROLLLEFT = &H2&
    DFCS_SCROLLRIGHT = &H3&
    DFCS_SCROLLCOMBOBOX = &H5&
    DFCS_SCROLLSIZEGRIP = &H8&
    DFCS_SCROLLSIZEGRIPRIGHT = &H10&
End Enum
#If False Then
    Private DFCS_SCROLLUP, DFCS_SCROLLDOWN, DFCS_SCROLLLEFT, DFCS_SCROLLRIGHT, DFCS_SCROLLCOMBOBOX, DFCS_SCROLLSIZEGRIP, DFCS_SCROLLSIZEGRIPRIGHT
#End If
Public Enum DFCButtonTypeFlags
'-- Button types:
    DFCS_BUTTONCHECK = &H0&
    DFCS_BUTTONRADIOIMAGE = &H1&
    DFCS_BUTTONRADIOMASK = &H2&
    DFCS_BUTTONRADIO = &H4&
    DFCS_BUTTON3STATE = &H8&
    DFCS_BUTTONPUSH = &H10&
End Enum
#If False Then
    Private DFCS_BUTTONCHECK, DFCS_BUTTONRADIOIMAGE, DFCS_BUTTONRADIOMASK, DFCS_BUTTONRADIO, DFCS_BUTTON3STATE, DFCS_BUTTONPUSH
#End If
Public Enum DFCStateTypeFlags
'-- Styles:
    DFCS_INACTIVE = &H100&
    DFCS_PUSHED = &H200&
    DFCS_CHECKED = &H400&
    '-- Win98/2000 only
    DFCS_TRANSPARENT = &H800&
    DFCS_HOT = &H1000&
    'End Win98/2000 only
    DFCS_ADJUSTRECT = &H2000&
    DFCS_FLAT = &H4000&
    DFCS_MONO = &H8000&
End Enum
#If False Then
    Private DFCS_INACTIVE, DFCS_PUSHED, DFCS_CHECKED, DFCS_TRANSPARENT, DFCS_HOT, DFCS_ADJUSTRECT, DFCS_FLAT, DFCS_MONO
#End If
Public Const CLR_INVALID                          As Integer = -1
Private Type PictDesc
    cbSizeofStruct                                    As Long
    picType                                           As Long
    hImage                                            As Long
    xExt                                              As Long
    yExt                                              As Long
End Type
Private Type Guid
    Data1                                             As Long
    Data2                                             As Integer
    Data3                                             As Integer
    Data4(0 To 7)                                     As Byte
End Type
'-- =======================================================================
'-- Image list Declares:
'-- =======================================================================
'-- Create/Destroy functions:
Private Const ILC_MASK                            As Long = 1
Private Const ILC_COLOR32                         As Long = &H20&
'-- Modification/deletion functions:
'-- Image information functions:
Private Type IMAGEINFO
    hBitmapImage                                      As Long
    hBitmapMask                                       As Long
    cPlanes                                           As Long
    cBitsPerPixel                                     As Long
    rcImage                                           As RECT
End Type
'-- Create a new icon based on an image list icon:
'-- Merge and move functions:

Private Type IMAGELISTDRAWPARAMS
    cbSize                                            As Long
    hIml                                              As Long
    i                                                 As Long
    hdcDst                                            As Long
    x                                                 As Long
    y                                                 As Long
    cX                                                As Long
    cY                                                As Long
    xBitmap                                           As Long
    '--        // x offest from the upperleft of bitmap
    yBitmap                                           As Long
    '--        // y offset from the upperleft of bitmap
    rgbBk                                             As Long
    rgbFg                                             As Long
    fStyle                                            As Long
    dwRop                                             As Long
End Type
''Private Const ILD_NORMAL                          As Integer = 0
Public Const ILD_TRANSPARENT                      As Integer = 1
Public Const ILD_BLEND25                          As Integer = 2
Public Const ILD_SELECTED                         As Integer = 4
Private Const FORMAT_MESSAGE_FROM_SYSTEM          As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS       As Long = &H200
Private m_hdcMono                                 As Long
Private m_hbmpMono                                As Long
Private m_hBmpOld                                 As Long
Public Type ICONINFO
    fIcon                                             As Long
    xHotspot                                          As Long
    yHotspot                                          As Long
    hbmMask                                           As Long
    hbmColor                                          As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal HWND As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
        ByVal nIndex As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
        ByVal nWidth As Long, _
        ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, _
        qrc As RECT, _
        ByVal edge As Long, _
        ByVal grfFlags As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
        ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
        ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
        ByVal nBkMode As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, _
        ByVal lpszExeFileName As String, _
        ByVal nIconIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
        ByVal lpsz As String, _
        ByVal un1 As Long, _
        ByVal n1 As Long, _
        ByVal n2 As Long, _
        ByVal un2 As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, _
        ByVal ptX As Long, _
        ByVal ptY As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, _
        ByVal xLeft As Long, _
        ByVal yTop As Long, _
        ByVal hIcon As Long, _
        ByVal cxWidth As Long, _
        ByVal cyWidth As Long, _
        ByVal istepIfAniCur As Long, _
        ByVal hbrFlickerFreeDraw As Long, _
        ByVal diFlags As Long) As Boolean
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
        ByVal X1 As Long, _
        ByVal y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
        ByVal x As Long, _
        ByVal y As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal HWND As Long, ByVal nCmdShow As Long) As Long

Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
        ByVal hBrush As Long, _
        ByVal lpDrawStateProc As Long, _
        ByVal lParam As Long, _
        ByVal wParam As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fuFlags As Long) As Long
Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Declare Function SelectObject Lib "gdi32" ( _
        ByVal hdc As Long, ByVal hObj As Long _
        ) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
        ByVal HPALETTE As Long, _
        pccolorref As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, _
        riid As Guid, _
        ByVal fPictureOwnsHandle As Long, _
        iPic As IPicture) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
        pSrc As Any, _
        ByVal ByteLen As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
        lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        Arguments As Long) As Long
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, _
        piconinfo As ICONINFO) As Long
Public Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long

Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal hdcDst As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal fStyle As Long) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, _
      ByVal i As Long, _
      ByVal diIgnore As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal hIml As Long, _
      ByVal i As Long, _
      prcImage As RECT) As Long



Public m_sCurrentSystemThemename          As String



Private Type OSVERSIONINFO
    dwVersionInfoSize                         As Long
    dwMajorVersion                            As Long
    dwMinorVersion                            As Long
    dwBuildNumber                             As Long
    dwPlatformId                              As Long
    szCSDVersion(0 To 127)                    As Byte
End Type
Private Const IDC_HAND                    As Long = 32649
Private Const IDC_ARROW                   As Long = 32512
Private Const VER_PLATFORM_WIN32_NT       As Integer = 2
Private Type TRIVERTEX
    x                                         As Long
    y                                         As Long
    Red                                       As Integer
    Green                                     As Integer
    Blue                                      As Integer
    Alpha                                     As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft                                 As Long
    LowerRight                                As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1                                   As Long
    Vertex2                                   As Long
    Vertex3                                   As Long
End Type
Public Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum
#If False Then
    Private GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V
#End If


Private m_bHandCursor                     As Boolean
Private m_bIsXp                           As Boolean
Public m_bIsNt                           As Boolean
Private m_bIs2000OrAbove                  As Boolean
Private m_bHasGradientAndTransparency     As Boolean
Public Declare Function IsWindow Lib "user32" (ByVal HWND As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal HWND As Long, _
        lpRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal HWND As Long, _
        ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
        ByVal dwMaxNameChars As Long, _
        ByVal pszColorBuff As Long, _
        ByVal cchMaxColorChars As Long, _
        ByVal pszSizeBuff As Long, _
        ByVal cchMaxSizeChars As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
        ByVal lpCursorName As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
        lpDeviceName As Any, _
        lpOutput As Any, _
        lpInitData As Any) As Long


Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal HWND As Long, _
        ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long








Public Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
        ByVal lpStr As String, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long
Public Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
        ByVal lpStr As Long, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, _
        lpRect As RECT) As Long




Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, _
        pVertex As TRIVERTEX, _
        ByVal dwNumVertex As Long, _
        pMesh As GRADIENT_RECT, _
        ByVal dwNumMesh As Long, _
        ByVal dwMode As Long) As Long


Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, _
        ByVal lHDC As Long, _
        ByVal iPartId As Long, _
        ByVal iStateId As Long, _
        pRect As RECT, _
        pClipRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal HWND As Long, _
        lpPoint As POINTAPI) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR  As Long

    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

Public Sub DrawOpenCloseGlyph(ByVal lngHwnd As Long, _
                              ByVal lHDC As Long, _
                              tTR As RECT, _
                              ByVal bCollapsed As Boolean)


    Dim tGR     As RECT
    Dim bDone   As Boolean
    Dim hTheme  As Long
    Dim hBr     As Long
    Dim hPen    As Long
    Dim hPenOld As Long
    Dim tJ      As POINTAPI

    LSet tGR = tTR

    With tGR
        .left = .left + 2
        .right = .left + 12
        .top = .top + 2    ' ((.bottom - .top) \ 2)
        .bottom = .top + 12
    End With    'tGR


    If IsXp Then
        hTheme = OpenThemeData(lngHwnd, StrPtr("TREEVIEW"))
        If Not (hTheme = 0) Then
            DrawThemeBackground hTheme, lHDC, 2, IIf(bCollapsed, 1, 2), tGR, tGR
            CloseThemeData hTheme
            bDone = True
        End If
    End If

    If Not (bDone) Then
        '-- Draw button border
        ' hBr = GetSysColorBrush(vbButtonFace And &H1F&)
        ' FillRect lHDC, tGR, hBr
        ' DeleteObject hBr
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbWindowText))
        hPenOld = SelectObject(lHDC, hPen)
        '

        mGraphics.UtilDrawBorderRectangle lHDC, vbBlack, tGR.left + 1, tGR.top + 1, 9, 9, False

        SelectObject lHDC, hPenOld
        DeleteObject hPen
        hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbWindowText))
        hPenOld = SelectObject(lHDC, hPen)
        '
        With tGR
            MoveToEx lHDC, .left + 3, .top + 5, tJ
            LineTo lHDC, .left + 8, .top + 5
        End With    'tGR
        If bCollapsed Then
            MoveToEx lHDC, tGR.left + 5, tGR.top + 3, tJ
            LineTo lHDC, tGR.left + 5, tGR.top + 8
        End If
        SelectObject lHDC, hPenOld
        DeleteObject hPen
    End If

End Sub

Private Sub GradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR         As GRADIENT_RECT
    Dim hBrush      As Long
    Dim lStartColor As Long
    Dim lEndColor   As Long

    'Dim lR As Long
    '-- Use GradientFill:
    If HasGradientAndTransparency Then
        lStartColor = TranslateColor(oStartColor)
        lEndColor = TranslateColor(oEndColor)
        setTriVertexColor tTV(0), lStartColor
        tTV(0).x = tR.left
        tTV(0).y = tR.top
        setTriVertexColor tTV(1), lEndColor
        tTV(1).x = tR.right
        tTV(1).y = tR.bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        '-- Fill with solid brush:
        hBrush = CreateSolidBrush(TranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub

Private Property Get HasGradientAndTransparency()


    HasGradientAndTransparency = m_bHasGradientAndTransparency

End Property

Public Property Get Is2000OrAbove() As Boolean


    Is2000OrAbove = m_bIs2000OrAbove

End Property

Public Property Get IsNt() As Boolean


    IsNt = m_bIsNt

End Property

Public Property Get IsXp() As Boolean

    IsXp = m_bIsXp

End Property

Private Sub setTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)

    Dim lRed   As Long
    Dim lGreen As Long
    Dim lBlue  As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        setTriVertexColorComponent .Red, lRed
        setTriVertexColorComponent .Green, lGreen
        setTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub

Public Sub UtilDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              ByVal lngLeft As Long, _
                              ByVal lngTop As Long, _
                              ByVal lngWidth As Long, _
                              ByVal lngHeight As Long, _
                              Optional ByVal horizontal As Boolean = False)

    Dim tR As RECT

    With tR
        .left = lngLeft
        .top = lngTop
        .right = lngLeft + lngWidth
        .bottom = lngTop + lngHeight
        '-- gradient fill vertical:
    End With    'tR
    GradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub

Public Sub UtilDrawBorderRectangle(ByVal lngHdc As Long, _
                                   ByVal lColor As Long, _
                                   ByVal lngLeft As Long, _
                                   ByVal lngTop As Long, _
                                   ByVal lngWidth As Long, _
                                   ByVal lngHeight As Long, _
                                   ByVal bInset As Boolean)


    Dim tJ      As POINTAPI
    Dim hPen    As Long
    Dim hPenOld As Long

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, lngLeft, lngTop + lngHeight - 1, tJ
    LineTo lngHdc, lngLeft, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop
    LineTo lngHdc, lngLeft + lngWidth - 1, lngTop + lngHeight - 1
    LineTo lngHdc, lngLeft, lngTop + lngHeight - 1
    SelectObject lngHdc, hPenOld
    DeleteObject hPen

End Sub

Public Sub UtilDrawText(ByVal lngHdc As Long, _
                        ByVal sCaption As String, _
                        ByVal lTextX As Long, _
                        ByVal lTextY As Long, _
                        ByVal lTextX1 As Long, _
                        ByVal lTextY1 As Long, _
                        ByVal bEnabled As Boolean, _
                        ByVal color As Long, _
                        ByVal bCentreHorizontal As Boolean, _
                        Optional RightAlign As Boolean = False)


    Dim rcText As RECT

    SetTextColor lngHdc, color
    'Dim lFlags As Long
    If Not bEnabled Then
        SetTextColor lngHdc, GetSysColor(vbGrayText And &H1F&)
    End If
    With rcText
        .left = lTextX
        .top = lTextY
        .right = lTextX1
        .bottom = lTextY1
    End With
    If m_bIsNt Then
        DrawTextW lngHdc, StrPtr(sCaption), -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    Else
        DrawTextA lngHdc, sCaption, -1, rcText, IIf(RightAlign, DT_RIGHT, DT_LEFT) Or DT_END_ELLIPSIS
    End If
    If Not bEnabled Then
        SetTextColor lngHdc, TranslateColor(vbWindowText)
    End If

End Sub

Public Sub UtilSetCursor(ByVal bHand As Boolean)

'-- Desc: Get the "Real" Hand Cursor

    If bHand Then
        SetCursor LoadCursor(0, IDC_HAND)
        m_bHandCursor = True
    Else
        SetCursor LoadCursor(0, IDC_ARROW)
        m_bHandCursor = False
    End If

End Sub

Public Sub VerInitialise()

    Dim tOSV As OSVERSIONINFO

    tOSV.dwVersionInfoSize = Len(tOSV)
    GetVersionEx tOSV
    m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If tOSV.dwMajorVersion > 5 Then
        m_bHasGradientAndTransparency = True
        m_bIsXp = True
        m_bIs2000OrAbove = True
    ElseIf (tOSV.dwMajorVersion = 5) Then
        m_bHasGradientAndTransparency = True
        m_bIs2000OrAbove = True
        If tOSV.dwMinorVersion >= 1 Then
            m_bIsXp = True
        End If
    ElseIf (tOSV.dwMajorVersion = 4) Then    '-- NT4 or 9x/ME/SE
        If tOSV.dwMinorVersion >= 10 Then
            m_bHasGradientAndTransparency = True
        End If
    Else    '-- Too old
    End If

End Sub


Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR) As Long


    Dim lCFrom As Long
    Dim lCTo   As Long
    Dim lCRetR As Long
    Dim lCRetG As Long
    Dim lCRetB As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
    If lCRetR > 255 Then
        lCRetR = 255
    ElseIf (lCRetR < 0) Then
        lCRetR = 0
    End If
    lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - ((lCFrom \ &H100) And &HFF&)) \ 2
    If lCRetG > 255 Then
        lCRetG = 255
    ElseIf (lCRetG < 0) Then
        lCRetG = 0
    End If
    lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - ((lCFrom \ &H10000) And &HFF&)) \ 2
    If lCRetB > 255 Then
        lCRetB = 255
    ElseIf (lCRetB < 0) Then
        lCRetB = 0
    End If
    BlendColor = RGB(lCRetR, lCRetG, lCRetB)

End Property

Private Sub HLSToRGB(ByVal h As Single, _
                     ByVal s As Single, _
                     ByVal l As Single, _
                     r As Long, _
                     g As Long, _
                     b As Long)


    Dim rR  As Single
    Dim rG  As Single
    Dim rB  As Single
    Dim min As Single
    Dim Max As Single

    If s = 0 Then
        '-- Achromatic case:
        rR = l
        rG = l
        rB = l
    Else
        '-- Chromatic case:
        '-- delta = Max-Min
        If l <= 0.5 Then
            's = (Max - Min) / (Max + Min)
            '-- Get Min value:
            min = l * (1 - s)
        Else
            's = (Max - Min) / (2 - Max - Min)
            '-- Get Min value:
            min = l - s * (1 - l)
        End If
        '-- Get the Max value:
        Max = 2 * l - min
        '-- Now depending on sector we can evaluate the h,l,s:
        If h < 1 Then
            rR = Max
            If h < 0 Then
                rG = min
                rB = rG - h * (Max - min)
            Else
                rB = min
                rG = h * (Max - min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If h < 2 Then
                rB = min
                rR = rB - (h - 2) * (Max - min)
            Else
                rR = min
                rB = (h - 2) * (Max - min) + rR
            End If
        Else
            rB = Max
            If h < 4 Then
                rR = min
                rG = rR - (h - 4) * (Max - min)
            Else
                rG = min
                rR = (h - 4) * (Max - min) + rG
            End If
        End If
    End If
    r = rR * 1.555
    g = rG * 1.555
    b = rB * 1.555

End Sub


Public Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long


    Static s_lLightColLast As Long

    Dim lC                 As Long
    Dim h                  As Single
    Dim s                  As Single
    Dim l                  As Single
    Dim lR                 As Long
    Dim lG                 As Long
    Dim lB                 As Long
    Static s_lColLast      As Long
    lC = TranslateColor(oColor)
    If lC <> s_lColLast Then
        s_lColLast = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If l > 0.99 Then
            l = l * 0.8
        Else
            l = l * 2
            If l > 1 Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        s_lLightColLast = RGB(lR, lG, lB)
    End If
    LighterColour = s_lLightColLast

End Property

Private Function Maximum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single

    If rR > rG Then
        If rR > rB Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If rB > rG Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If

End Function

Private Function Minimum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single

    If rR < rG Then
        If rR < rB Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If rB < rG Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If

End Function


Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object

    Dim oTemp As Object

    '-- Turn the pointer into an illegal, uncounted interface
    CopyMemory oTemp, lPtr, 4
    '-- Do NOT hit the End button here! You will crash!
    '-- Assign to legal reference
    Set ObjectFromPtr = oTemp
    '-- Still do NOT hit the End button here! You will still crash!
    '-- Destroy the illegal reference
    CopyMemory oTemp, 0&, 4
    '-- OK, hit the End button if you must--you'll probably still crash,
    '-- but it will be because of the subclass, not the uncounted reference

End Property

Private Function ResizeIconImage(hImg As Long) As Long


    Dim icoInfo    As ICONINFO
    Dim newICOinfo As ICONINFO
    Dim icoBMPinfo As BITMAP

    GetIconInfo hImg, icoInfo
    If ResizeIconImage Then
        DestroyIcon ResizeIconImage
    End If
    '-- start a new icon structure
    CopyMemory newICOinfo, icoInfo, Len(icoInfo)
    '-- get the icon dimensions from the bitmap portion of the icon
    GetGDIObject icoInfo.hbmColor, Len(icoBMPinfo), icoBMPinfo
    ResizeIconImage = CreateIconIndirect(newICOinfo)
    DeleteObject newICOinfo.hbmMask
    DeleteObject newICOinfo.hbmColor

End Function

Private Sub RGBToHLS(ByVal r As Long, _
                     ByVal g As Long, _
                     ByVal b As Long, _
                     h As Single, _
                     s As Single, _
                     l As Single)


    Dim Max   As Single
    Dim min   As Single
    Dim delta As Single
    Dim rR    As Single
    Dim rG    As Single
    Dim rB    As Single

    rR = r / 255
    rG = g / 255
    rB = b / 255
    '{Given: rgb each in [0,1].
    '-- Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    min = Minimum(rR, rG, rB)
    l = (Max + min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = min Then
        'begin {Acrhomatic case}
        s = 0
        h = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
        '{First calculate the saturation.}
        If l <= 0.5 Then
            s = (Max - min) / (Max + min)
        Else
            s = (Max - min) / (2 - Max - min)
        End If
        '{Next calculate the hue.}
        delta = Max - min
        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta    '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta    '{Resulting color is between magenta and cyan}
        End If
        'end {Chromatic Case}
    End If

End Sub

Public Property Get SlightlyLighterColour(ByVal oColor As OLE_COLOR) As Long


    Static s_lLightColLast As Long

    Dim lC                 As Long
    Dim h                  As Single
    Dim s                  As Single
    Dim l                  As Single
    Dim lR                 As Long
    Dim lG                 As Long
    Dim lB                 As Long
    Static s_lColLast      As Long
    lC = TranslateColor(oColor)
    If lC <> s_lColLast Then
        s_lColLast = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
        If l > 0.99 Then
            l = l * 0.95
        Else
            l = l * 2
            If l > 1 Then
                l = 1
            End If
        End If
        HLSToRGB h, s, l, lR, lG, lB
        s_lLightColLast = RGB(lR, lG, lB)
    End If
    SlightlyLighterColour = s_lLightColLast

End Property

Public Sub TileArea(ByVal hdcTo As Long, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal lngWidth As Long, _
                    ByVal lngHeight As Long, _
                    ByVal hDcSrc As Long, _
                    ByVal SrcWidth As Long, _
                    ByVal SrcHeight As Long, _
                    ByVal lOffsetY As Long)


    Dim lSrcX           As Long
    Dim lSrcY           As Long
    Dim lSrcStartX      As Long
    Dim lSrcStartY      As Long
    Dim lSrcStartWidth  As Long
    Dim lSrcStartHeight As Long
    Dim lDstX           As Long
    Dim lDstY           As Long
    Dim lDstWidth       As Long
    Dim lDstHeight      As Long

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    lDstY = y
    lDstHeight = lSrcStartHeight
    Do While lDstY < (y + lngHeight)
        If (lDstY + lDstHeight) > (y + lngHeight) Then
            lDstHeight = y + lngHeight - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + lngWidth)
            If (lDstX + lDstWidth) > (x + lngWidth) Then
                lDstWidth = x + lngWidth - lDstX
                If lDstWidth = 0 Then
                    lDstWidth = 4
                End If
            End If
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop

End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

'-- Convert Automation color to Windows color

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function





