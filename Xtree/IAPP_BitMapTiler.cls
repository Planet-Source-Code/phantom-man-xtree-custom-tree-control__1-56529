VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_BitMapTiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'TreeProps
'//---------------------------------------------------------------------------------------
'-- Module    : IAPP_BitmapTiler
'-- DateTime  : 26/07/2004 13:47
'-- Author    : Gary Noble
'-- Purpose   : Tiles a picture over the specified area in a DC
'-- Assumes   :
'//---------------------------------------------------------------------------------------
'-- History   : Initial Implementation    Gary Noble  26/07/2004
'//---------------------------------------------------------------------------------------
Option Explicit
Private m_lHdc                  As Long
Private m_lHBmp                 As Long
Private m_lHBmpOld              As Long
Private m_lhPalOld              As Long
Private m_pic                   As StdPicture
Private m_sFileName             As String
Private m_lXOriginOffset        As Long
Private m_lYOriginOffset        As Long
Private m_lBitmapW              As Long
Private m_lBitmapH              As Long
Private Const IAPP_BitmapTilerErrorBase    As Integer = 5600
Private Type BITMAP
    bmType                          As Long
    bmWidth                         As Long
    bmHeight                        As Long
    bmWidthBytes                    As Long
    bmPlanes                        As Integer
    bmBitsPixel                     As Integer
    bmBits                          As Long
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
        ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
        ByVal nCount As Long, _
        lpObject As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal HWND As Long, _
        ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public Property Get BitmapHeight() As Long


    BitmapHeight = m_lBitmapH

End Property

Public Property Get BitmapWidth() As Long


    BitmapWidth = m_lBitmapW

End Property

Private Sub Class_Terminate()

'-- Ensure all GDI objects are freed:

    pClearUp
    '-- Clear up the picture:
    Set m_pic = Nothing

End Sub

Public Property Get Filename() As String

    Filename = m_sFileName

End Property

Public Property Let Filename(ByVal sFileName As String)

'-- Load a picture from a file:

    If m_sFileName <> sFileName Then
        pClearUp
        If pbLoadPicture(sFileName) Then
            m_sFileName = sFileName
        End If
    End If

End Property

Public Property Get Palette() As StdPicture


    Set Palette = m_pic

End Property

Private Function pbEnsurePicture() As Boolean

    On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture
        If Err.Number <> 0 Then
            pErr 3, "Unable to allocate memory for picture object."
            pbEnsurePicture = False
        Else
        End If
    End If
    On Error GoTo 0

End Function

Private Function pbGetBitmapIntoDC() As Boolean

    Dim tB           As BITMAP
    Dim lHDC         As Long
    Dim lHwnd        As Long
    Dim lHDCTemp     As Long
    Dim lHBmpTempOld As Long

    '-- Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    lHDCTemp = CreateCompatibleDC(lHDC)
    If m_lHdc <> 0 Then
        '-- Get size of bitmap:
        GetObjectAPI m_pic.Handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.Handle)
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        SelectObject lHDCTemp, lHBmpTempOld
        DeleteDC lHDCTemp
        If m_lHBmpOld <> 0 Then
            pbGetBitmapIntoDC = True
            If LenB(m_sFileName) = 0 Then
                m_sFileName = "PICTURE"
            End If
        Else
            pClearUp
            pErr 2, "Unable to select bitmap into DC"
        End If
    Else
        pErr 1, "Unable to create compatible DC"
    End If
    ReleaseDC lHwnd, lHDC

End Function

Private Function pbLoadPicture(sFile As String) As Boolean

    If pbEnsurePicture() Then
        If pbLoadPictureFromFile(sFile) Then
            pbLoadPicture = pbGetBitmapIntoDC()
        End If
    End If

End Function

Private Function pbLoadPictureFromFile(ByVal sFile As String) As Boolean


    On Error Resume Next
    Set m_pic = LoadPicture(sFile)
    If Err.Number <> 0 Then
        pErr 0, "Load Picture Failed: " & Err.Description
    Else
        pbLoadPictureFromFile = True
    End If
    On Error GoTo 0

End Function

Private Sub pClearUp()

'-- Clear reference to the filename:

    m_sFileName = ""
    '-- If we have a DC, then clear up:
    If m_lHdc <> 0 Then
        '-- Select the bitmap out of DC:
        If m_lHBmpOld <> 0 Then
            SelectObject m_lHdc, m_lHBmpOld
            '-- The original bitmap does not have to deleted because it is owned by m_pic
        End If
        If m_lHBmp <> 0 Then
            DeleteObject m_lHBmp
        End If
        '-- Remove the DC:
        DeleteDC m_lHdc
    End If

End Sub

Private Sub pErr(ByVal lNumber As Long, _
                 ByVal smsg As String)


    MsgBox "Error: " & smsg & ", " & lNumber, vbExclamation

End Sub

Public Property Get Picture() As StdPicture

    Set Picture = m_pic

End Property

Public Property Let Picture(oPic As StdPicture)

'-- Load a picture from a StdPicture object:

    pClearUp
    If Not oPic Is Nothing Then
        If pbEnsurePicture() Then
            Set m_pic = oPic
            If Err.Number = 0 Then
                pbGetBitmapIntoDC
            End If
        End If
    End If

End Property

Public Property Get PicturehDC() As Long


    PicturehDC = m_lHdc

End Property

Public Sub TileArea(ByRef lngHdc As Long, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal lngWidth As Long, _
                    ByVal lngHeight As Long)


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

    lSrcStartX = ((x + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
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
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt lngHdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop

End Sub

Public Property Get XOriginOffset() As Long

    XOriginOffset = m_lXOriginOffset

End Property

Public Property Let XOriginOffset(ByVal lPixels As Long)

    m_lXOriginOffset = lPixels

End Property

Public Property Get YOriginOffset() As Long

    YOriginOffset = m_lYOriginOffset

End Property

':) Roja's VB Code Fixer V1.0.66 (26/07/2004 10:49:28) 42 + 268 = 310 Lines Thanks Ulli for inspiration and lots of code.
Public Property Let YOriginOffset(ByVal lPiYels As Long)

    m_lYOriginOffset = lPiYels

End Property




