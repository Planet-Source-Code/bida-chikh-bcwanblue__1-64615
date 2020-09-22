VERSION 5.00
Begin VB.UserControl CButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.Timer tm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "CButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const bDefPCA = 0
Const bDefBS = 0

Dim RetVal As Long
Dim LTVal As RECT
Dim XYVal As POINT_TYPE
Dim bWidth, bHeight
Dim CPosX, CPosY
Dim bLeft, bTop

Dim cOut As Boolean
Dim cPress As Boolean
Dim kPress As Boolean
Dim cIn As Boolean

Enum PicCapAlignConstants
  [AlignNone]
  [AlignCenter]
  [AlignLeft]
  [AlignRight]
  [AligneTop]
  [AlignBottom]
End Enum
Enum BordStyleConstants
  [Flat]
  [Fixed Single]
End Enum

Dim bLColor As OLE_COLOR
Dim bSColor As OLE_COLOR
Dim bBColor As OLE_COLOR
Dim bFColor As OLE_COLOR
Dim bFSize As Integer
Dim WithEvents bFont As StdFont
Attribute bFont.VB_VarHelpID = -1
Dim bCaption As String
Dim bhWnd As Long
Dim bhDC As Long
Dim bPicture As StdPicture
Dim bMaskColor As OLE_COLOR
Dim bUMColor As Boolean
Dim bPCAlign As PicCapAlignConstants
Dim bCAlign
Dim bBrdStyle As BordStyleConstants
Dim bBrdColor As OLE_COLOR
Dim bHovColor As OLE_COLOR
Dim bHovPicture As StdPicture
Dim bTmpPicture As StdPicture
Dim bCM
Dim bGoTFocus As Boolean
Dim bEnabled As Boolean


Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut()
Event MouseIn()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub DrawButton()
  UserControl.Cls
  If (cPress = True) Or (kPress = True) Then
    bLColor = &H808080
    bSColor = &HFFFFFF
  End If
  If ((cPress = False) And (kPress = False)) Or _
     ((cOut = True) And (cPress = True)) Then
    bLColor = &HFFFFFF
    bSColor = &H808080
  End If
  If ((cOut = True) And (cPress = False)) And kPress = False Then
    If bBrdStyle = Flat Then
      bLColor = bBColor
      bSColor = bBColor
    Else
      bLColor = bBrdColor
      bSColor = bBrdColor
    End If
  End If

  Call DrawPicture
  Call DrawCaption

  UserControl.Line (0, 0)-(bWidth, 0), bLColor
  UserControl.Line (0, 0)-(0, bHeight), bLColor
  UserControl.Line (bWidth - 1, 0)-(bWidth - 1, bHeight), bSColor
  UserControl.Line (0, bHeight - 1)-(bWidth - 1, bHeight - 1), bSColor
End Sub

Private Sub DrawCaption()
  If bTmpPicture Is Nothing Then
    bFSize = bFont.Size
    Select Case bPCAlign
      Case 0, 1
        CPosX = (bWidth / 2)
        CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
        bCAlign = TA_CENTER
      Case 2
        CPosX = 4
        CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
        bCAlign = TA_LEFT
      Case 3
        CPosX = bWidth - 5
        CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
        bCAlign = TA_RIGHT
      Case 4
        CPosX = (bWidth / 2)
        CPosY = 4
        bCAlign = TA_TOP Or TA_CENTER
      Case 5
        CPosX = (bWidth / 2)
        CPosY = bHeight - bFSize
        bCAlign = TA_BOTTOM Or TA_CENTER
    End Select
  End If

  If bEnabled = False Then
    UserControl.ForeColor = &HFFFFFF
    RetVal = SetTextAlign(hdc, bCAlign)
    RetVal = TextOut(hdc, CPosX + 1, CPosY + 1, bCaption, Len(bCaption))
    UserControl.ForeColor = &H808080
    RetVal = SetTextAlign(hdc, bCAlign)
    RetVal = TextOut(hdc, CPosX, CPosY, bCaption, Len(bCaption))
    Exit Sub
  End If

  'UserControl.ForeColor = bFColor

  RetVal = SetTextAlign(hdc, bCAlign)
  RetVal = TextOut(hdc, CPosX + bCM, CPosY + bCM, bCaption, Len(bCaption))
End Sub

Private Sub DrawPicture()
  Dim ighDC As Long
  Dim ighWnd As Long
  Dim igWidth, igHeight
  Dim PosX, PosY

  If bTmpPicture Is Nothing Then Exit Sub
  igWidth = Int(bTmpPicture.Width / 26.455)
  igHeight = Int(bTmpPicture.Height / 26.455)

  bFSize = bFont.Size
  Select Case bPCAlign
    Case 0
      'bPicture = LoadPicture("")
      CPosX = (bWidth / 2)
      CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
      bCAlign = TA_CENTER
      Exit Sub
    Case 1
      PosX = Int((bWidth - igWidth) / 2)
      PosY = Int((bHeight - igHeight) / 2)
      CPosX = (bWidth / 2)
      CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
      bCAlign = TA_CENTER
    Case 2
      PosX = 4
      PosY = Int((bHeight - igHeight) / 2)
      CPosX = (igWidth + 8)
      CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
      bCAlign = TA_LEFT
    Case 3
      PosX = bWidth - (igWidth + 5)
      PosY = Int((bHeight - igHeight) / 2)
      CPosX = PosX - 4
      CPosY = Int((bHeight / 2) - (bFSize - (bFSize / 3.7)))
      bCAlign = TA_RIGHT
    Case 4
      PosX = Int((bWidth - igWidth) / 2)
      PosY = 4
      CPosX = (bWidth / 2)
      CPosY = igHeight + 4
      bCAlign = TA_CENTER Or TA_TOP
    Case 5
      PosX = Int((bWidth - igWidth) / 2)
      PosY = bHeight - (igHeight + 5)
      CPosX = (bWidth / 2)
      CPosY = PosY
      bCAlign = TA_CENTER Or TA_BOTTOM
  End Select

  ighWnd = bTmpPicture.Handle
  ighDC = CreateCompatibleDC(hdc)
  Call SelectObject(ighDC, ighWnd)

  If bUMColor Then
    Select Case bTmpPicture.Type
      Case vbPicTypeBitmap
        Call TransparentBlt(hdc, PosX + bCM, PosY + bCM, _
             igWidth, igHeight, ighDC, 0, 0, _
             igWidth, igHeight, bMaskColor)
      Case vbPicTypeIcon
        Call DrawIconEx(hdc, PosX + bCM, PosY + bCM, ighWnd, igWidth, _
             igHeight, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE)
      Case vbPicTypeMetafile, vbPicTypeEMetafile
        PaintPicture bTmpPicture, PosX + bCM, PosY + bCM '''
    End Select
  Else
    Select Case bTmpPicture.Type
      Case vbPicTypeIcon, vbPicTypeMetafile, vbPicTypeEMetafile
        PaintPicture bTmpPicture, PosX + bCM, PosY + bCM ''
      Case Else
        Call BitBlt(hdc, PosX + bCM, PosY + bCM, _
             igWidth, igHeight, _
             ighDC, 0, 0, SRCCOPY)
    End Select
  End If
  Call DeleteDC(ighDC)
End Sub

Private Sub bFont_FontChanged(ByVal PropertyName As String)
  'Set UserControl.Font = bFont
  'Call DrawButton
End Sub

Private Sub tm_Timer()
  RetVal = GetWindowRect(hwnd, LTVal)
  RetVal = GetCursorPos(XYVal)

  If XYVal.X < LTVal.left Or XYVal.X > LTVal.right Or _
     XYVal.Y < LTVal.top Or XYVal.Y > LTVal.bottom Then
    cOut = True
    cIn = False
    UserControl.ForeColor = bFColor
    If bPicture Is Nothing Then
      Set bTmpPicture = Nothing
    Else
      Set bTmpPicture = bPicture
    End If
    Call DrawButton
    If cPress = True Then
      Exit Sub
    Else
      RaiseEvent MouseOut
      tm.Enabled = False
    End If
  Else
    cOut = False
    UserControl.ForeColor = bHovColor
    If bHovPicture Is Nothing Then
      Set bTmpPicture = Nothing
      If bPicture Is Nothing Then
        Set bTmpPicture = Nothing
      Else
        Set bTmpPicture = bPicture
      End If
    Else
      Set bTmpPicture = bHovPicture
    End If
    If cIn = False Then
      Call DrawButton
      RaiseEvent MouseIn
      cIn = True
    End If
  End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  'enabled prop

  If KeyAscii = 13 Or KeyAscii = 27 Then RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Call DrawButton
End Sub

Private Sub UserControl_DblClick()
  Call SetCapture(hwnd)
  UserControl_MouseDown 1, 0, 0, 0
End Sub

Private Sub UserControl_EnterFocus()
  bGoTFocus = True
  cIn = False
  Call DrawButton
End Sub

Private Sub UserControl_ExitFocus()
  bGoTFocus = False
  kPress = False
  cPress = False
  cOut = True
  cIn = False
  Call DrawButton
End Sub

Private Sub UserControl_Initialize()
  kPress = False
  cPress = False
  cOut = True
  cIn = False
  Set bFont = New StdFont
End Sub

Private Sub UserControl_InitProperties()
  bLColor = &HFFFFFF
  bSColor = &H808080
  bBColor = &H8000000F
  bCaption = Ambient.DisplayName
  bFSize = UserControl.Font.Size
  Set UserControl.Font = Ambient.Font
  Set bFont = Ambient.Font
  Set bPicture = LoadPicture("")
  Set bHovPicture = LoadPicture("")
  Set bTmpPicture = LoadPicture("")
  bPCAlign = AlignNone
  bBrdStyle = Flat
  bFColor = Ambient.ForeColor
  bBrdColor = &HE0E0E0
  bEnabled = True
  bHovColor = Ambient.ForeColor
  UserControl.BackColor = bBColor
  UserControl.ForeColor = bFColor
  UserControl.Enabled = bEnabled
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeySpace
      kPress = True
      bCM = 1
      RaiseEvent Click
      Call DrawButton
  End Select
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeySpace
      kPress = False
      bCM = -1
      Call DrawButton
  End Select
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> vbRightButton Then
    cPress = True
    bCM = 1
    Call DrawButton
  End If
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If cOut = True Then tm.Enabled = True
  RaiseEvent MouseOver(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  cPress = False
  If Button <> vbRightButton Then
    bCM = -1
    Call DrawButton
    If (X >= 0 And Y >= 0) And (X <= bWidth And Y <= bHeight) Then RaiseEvent Click
  End If
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
  bFSize = bFont.Size
  bWidth = UserControl.ScaleWidth
  bHeight = UserControl.ScaleHeight
  If bHeight < bFSize Then
    UserControl.Height = ((bFSize + (bFSize / 2.5)) * 15)
    bHeight = UserControl.ScaleHeight
  End If
  Call DrawButton
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = bBColor
End Property

Public Property Let BackColor(ByVal bColor As OLE_COLOR)
  bBColor = bColor
  UserControl.BackColor = bBColor
  PropertyChanged "BackColor"
  Call DrawButton
End Property
Public Property Get Caption() As String
  Caption = bCaption
End Property

Public Property Let Caption(ByVal cap As String)
  bCaption = cap
  PropertyChanged "Caption"
  Call DrawButton
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = bFColor
End Property

Public Property Let ForeColor(ByVal fcolor As OLE_COLOR)
  bFColor = fcolor
  UserControl.ForeColor = bFColor
  PropertyChanged "ForeColor"
  Call DrawButton
End Property

Public Property Get Font() As Font
  Set Font = bFont
End Property

Public Property Set Font(ByVal fnt As Font)
  With bFont
   .Name = fnt.Name
   .Size = fnt.Size
   .Bold = fnt.Bold
   .Italic = fnt.Italic
   .Underline = fnt.Underline
   .Strikethrough = fnt.Strikethrough
  End With
  Set UserControl.Font = bFont
  UserControl_Resize
  PropertyChanged "Font"
  Call DrawButton
End Property

Public Property Get MaskColor() As OLE_COLOR
  MaskColor = bMaskColor
End Property

Public Property Let MaskColor(ByVal mc As OLE_COLOR)
  bMaskColor = mc
  PropertyChanged "MaskColor"
  Call DrawButton
End Property

Public Property Get UseMaskColor() As Boolean
  UseMaskColor = bUMColor
End Property

Public Property Let UseMaskColor(ByVal umc As Boolean)
  bUMColor = umc
  PropertyChanged "UseMaskColor"
  Call DrawButton
End Property

Public Property Get Picture() As StdPicture
  Set Picture = bPicture
End Property

Public Property Set Picture(ByVal pic As StdPicture)
  Set bPicture = pic
  Set bTmpPicture = bPicture
  PropertyChanged "Picture"
  Call DrawButton
End Property

Public Property Get PicCapalign() As PicCapAlignConstants
  PicCapalign = bPCAlign
End Property

Public Property Let PicCapalign(ByVal pcalign As PicCapAlignConstants)
  bPCAlign = pcalign
  PropertyChanged "PicCapAlign"
  Call DrawButton
End Property

Public Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property

Public Property Get BorderStyle() As BordStyleConstants
  BorderStyle = bBrdStyle
End Property

Public Property Let BorderStyle(ByVal brdstyle As BordStyleConstants)
  bBrdStyle = brdstyle
  PropertyChanged "BorderStyle"
  Call DrawButton
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = bBrdColor
End Property

Public Property Let BorderColor(ByVal brdcolor As OLE_COLOR)
  bBrdColor = brdcolor
  PropertyChanged "BorderColor"
  Call DrawButton
End Property

Public Property Get HoverColor() As OLE_COLOR
  HoverColor = bHovColor
End Property

Public Property Let HoverColor(ByVal hovcolor As OLE_COLOR)
  bHovColor = hovcolor
  PropertyChanged "HoverColor"
End Property

Public Property Get HoverPicture() As StdPicture
  Set HoverPicture = bHovPicture
End Property

Public Property Set HoverPicture(ByVal hovpic As StdPicture)
  Set bHovPicture = hovpic
  PropertyChanged "HoverPicture"
  Call DrawButton
End Property

Public Property Get Enabled() As Boolean
  Enabled = bEnabled
End Property

Public Property Let Enabled(ByVal ena As Boolean)
  bEnabled = ena
  PropertyChanged "Enabled"
  Call DrawButton
End Property

Public Property Get MousePointer() As MousePointerConstants
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal mp As MousePointerConstants)
  UserControl.MousePointer() = mp
  PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal mi As StdPicture)
  If mi.Type <> vbPicTypeIcon Then
    MsgBox "Invalid MouseIcon File!", vbCritical, "CButton"
    Exit Property
  End If
  Set UserControl.MouseIcon = mi
  PropertyChanged "MouseIcon"
End Property

Public Sub Refresh()
  Call DrawButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", bBColor, &H8000000F)
  Call PropBag.WriteProperty("Caption", bCaption, Ambient.DisplayName)
  Call PropBag.WriteProperty("ForeColor", bFColor, Ambient.ForeColor)
  Call PropBag.WriteProperty("Font", bFont, Ambient.Font)
  Call PropBag.WriteProperty("Picture", bPicture, Nothing)
  Call PropBag.WriteProperty("UseMaskColor", bUMColor, False)
  Call PropBag.WriteProperty("MaskColor", bMaskColor, &H8000000F)
  Call PropBag.WriteProperty("PicCapAlign", bPCAlign, bDefPCA)
  Call PropBag.WriteProperty("BorderStyle", bBrdStyle, bDefBS)
  Call PropBag.WriteProperty("BorderColor", bBrdColor, &HE0E0E0)
  Call PropBag.WriteProperty("HoverColor", bHovColor, bFColor)
  Call PropBag.WriteProperty("HoverPicture", bHovPicture, Nothing)
  Call PropBag.WriteProperty("Enabled", bEnabled, True)
  Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  bBColor = PropBag.ReadProperty("BackColor", &H8000000F)
  bCaption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
  bFColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
  Set bFont = PropBag.ReadProperty("Font", Ambient.Font)
  Set bPicture = PropBag.ReadProperty("Picture", Nothing)
  bUMColor = PropBag.ReadProperty("UseMaskColor", bUMColor)
  bMaskColor = PropBag.ReadProperty("MaskColor", &H8000000F)
  bPCAlign = PropBag.ReadProperty("PicCapAlign", bDefPCA)
  bBrdStyle = PropBag.ReadProperty("BorderStyle", bDefBS)
  bBrdColor = PropBag.ReadProperty("BorderColor", &HE0E0E0)
  bHovColor = PropBag.ReadProperty("HoverColor", bFColor)
  Set bHovPicture = PropBag.ReadProperty("HoverPicture", Nothing)
  bEnabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)

  UserControl.BackColor = bBColor
  UserControl.ForeColor = bFColor
  UserControl.Enabled = bEnabled
  Set UserControl.Font = bFont
  Set bTmpPicture = bPicture

  Call DrawButton
End Sub

