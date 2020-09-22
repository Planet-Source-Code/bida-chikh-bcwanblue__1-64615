Attribute VB_Name = "APIs"
Option Explicit

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32.dll" () As Long
Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDst As Long, ByVal nXOriginDst As Long, ByVal nYOriginDst As Long, ByVal nWidthDst As Long, ByVal nHeightDst As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Const BLACKNESS = &H42
Public Const DSTINVERT = &H550009
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const PATCOPY = &HF00021
Public Const PATINVERT = &H5A0049
Public Const PATPAINT = &HFB0A09
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const WHITENESS = &HFF0062

Public Const TA_BASELINE = 24
Public Const TA_BOTTOM = 8
Public Const TA_CENTER = 6
Public Const TA_LEFT = 0
Public Const TA_NOUPDATECP = 0
Public Const TA_RIGHT = 2
Public Const TA_RTLREADING = 256
Public Const TA_TOP = 0
Public Const TA_UPDATECP = 1

Type POINT_TYPE
  X As Long
  Y As Long
End Type

Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type
