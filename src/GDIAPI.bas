Attribute VB_Name = "GDIAPI"
Option Explicit
Type point
        x As Long
        y As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Type hBitmap
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type


Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type

Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Single 'Ushort value
   Green As Single 'Ushort value
   Blue As Single 'ushort value
   Alpha As Integer 'ushort
End Type

Public Const Blackness = 66
Public Const NotSrcErase = 1114278
Public Const NotSrcCopy = 3342344
Public Const SrcErase = 4457256
Public Const DstInvert = 5570569
Public Const PatInvert = 5898313
Public Const SrcInvert = 6684742
Public Const SrcAnd = 8913094
Public Const MergePaint = 12255782
Public Const MergeCopy = 12583114
Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = origen
Public Const ScrPaint = 15597702
Public Const PatCopy = 15728673
Public Const PatPaint = 16452105
Public Const Whiteness = 16711778

Global Const KEY_LBUTTON = &H1
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F
Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F

Public Const IMAGE_BITMAP As Long = 0
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const PI = 3.141593
Public Const PI_180 = 3.141593 / 180

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0   '  tabla de color en RGB (rojo-verde-azul)
Public Const DIB_PAL_COLORS = 1   '  tabla de color en los índices de la paleta

Public Const GL_FUNC_ADD = 32774
Public Const GL_FUNC_SUBTRACT = 32778
Public Const GL_FUNC_REVERSE_SUBTRACT = 32779

Const HORZRES As Long = 8
Const VERTRES As Long = 10
Const OBJ_BITMAP = 7
Const GRADIENT_FILL_TRIANGLE = &H2
Public Const HALFTONE = 4

Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" _
(ByRef saArray() As Any) As Long

Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal TopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Declare Function CreatePen Lib "gdi32" (ByVal style As Long, ByVal width As Long, ByVal color As Long) As Long
Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long _
)

Declare Function MaskBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nwidth As Long, ByVal nheight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long, ByVal dwRop As Long) As Long


Declare Function GetPixel Lib "gdi32" (ByVal hdcOrigin As Long, ByVal x As Long, ByVal y As Long) As Long

Declare Function SetPixel Lib "gdi32" (ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
                                     ByVal nwidth As Long, ByVal nheight As Long, _
                                     ByVal hSrcDC As Long, ByVal xSrc As Long, _
                                     ByVal ySrc As Long, ByVal dwRop As Long) As Long
                                     
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal Xend As Long, _
                                     ByVal Yend As Long) As Long
Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, _
                                                      ByVal CurX As Long, _
                                                      ByVal CurY As Long, _
                                                      vbnull) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                           ByVal nwidth As Long, _
                                           ByVal nheight As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                        ByVal iBkMode As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                         ByVal crColor As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long


Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public memDC As Long, membm As Long, memdc2 As Long, res As Long
Public cap As String, percent As Long


'Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal BITMAP As Long, ByRef lpDimension As Point) As Long
Declare Function GetTickCount Lib "KERNEL32" () As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nwidth As Long, ByVal nheight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal _
   uObjectType As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal _
   hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Declare Function Polygon Lib "gdi32" ( _
'    ByVal hdc As Long, _
'    ByRef points As Point, _
'    ByVal PointNum As Long) As Long
'Declare Function CreateBrushIndirect Lib "gdi32" (ByRef lpLogBrush As LOGBRUSH) As Integer
Declare Function GradientFill Lib "msimg32" _
  (ByVal hdc As Long, _
   pVertex As Any, _
   ByVal dwNumVertex As Long, _
   pMesh As Any, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
  ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Declare Function GetCommandLine Lib "KERNEL32" Alias "GetCommandLineA" () As Long
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDst As Any, _
    pSrc As Any, ByVal ByteLen As Long)
Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function GetViewportExtEx Lib "gdi32" (ByVal hdc As Long, lpSize As SIZE) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" ( _
Destination As Any, _
ByVal Length As Long _
)

Declare Sub gluPerspective Lib "glu32.dll" (ByVal fovy#, ByVal aspect#, ByVal zNear#, ByVal zFar#)
Declare Function gluProject Lib "glu32.dll" (ByVal objx As Double, ByVal objy As Double, ByVal objz As Double, modelMatrix As Any, projMatrix As Any, viewport As Any, winx As Double, winy As Double, winz As Double) As Long
Declare Function gluUnProject Lib "glu32.dll" (ByVal winx As Double, ByVal winy As Double, ByVal winz As Double, modelMatrix As Any, projMatrix As Any, viewport As Any, objx As Double, objy As Double, objz As Double) As Long
Declare Sub gluPickMatrix Lib "glu32.dll" (ByVal x As Double, ByVal y As Double, ByVal nwidth As Double, ByVal nheight As Double, viewport As Any)

Function creaDC(ByVal dc As Long, ByVal x As Long, ByVal y As Long) As Long
Dim hBitmap As Long, hdc As Long, tipo As Integer, error As String
error = "Error de gráficos!"
tipo = 0 + 0 + 16
Do
    hdc = CreateCompatibleDC(dc)
Loop While hdc < 1
'If hDC < 1 Then
'    MsgBox "No se puede crear el contexto de dispositivo", tipo, error
'    creaDC = 0
'    Exit Function
'End If
hBitmap = CreateCompatibleBitmap(dc, x, y)
If hBitmap = 0 Then
    MsgBox "No se puede crear el bitmap", tipo, error
    DeleteDC hdc
    Exit Function
End If
SelectObject hdc, hBitmap
creaDC = hdc
DeleteObject hBitmap
BitBlt creaDC, 0, 0, x, y, creaDC, 0, 0, Blackness
End Function
Sub GetHeaderBitmapInfo(ByVal hdc As Long, ByVal hbmp As Long, ByRef BMPinfo As BITMAPINFO)
    BMPinfo.bmiHeader.biSize = 40
    
    GetDIBits hdc, hbmp, 0, 0, ByVal 0&, BMPinfo, DIB_RGB_COLORS
End Sub
'Uses the 3-call method to GetDIBits for retriving the bitmap data (including original pallete)
Sub GetAllBitmapData(ByVal hdc As Long, ByVal hbmp As Long, ByRef BMPData() As Byte, ByRef BMPinfo As BITMAPINFO)
    
    GetHeaderBitmapInfo hdc, hbmp, BMPinfo
    
    If (BMPinfo.bmiHeader.biBitCount <= 8) Then
        Dim OldUsed As Long
        OldUsed = BMPinfo.bmiHeader.biClrUsed ' Read bitmap palette
        Call GetDIBits(hdc, hbmp, 0, 0, ByVal 0&, BMPinfo, 0)
        BMPinfo.bmiHeader.biClrUsed = OldUsed
    End If
    
    With BMPinfo.bmiHeader ' Allocate data array
        ReDim BMPData((((((.biWidth * .biBitCount) + 31) And _
            &HFFFFFFE0) \ &H20) * .biHeight) * 4 - 1) As Byte
    End With
    
    ' Read image data
    Call GetDIBits(hdc, hbmp, 0, _
        BMPinfo.bmiHeader.biHeight, BMPData(0), BMPinfo, 0)
End Sub

