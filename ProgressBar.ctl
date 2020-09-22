VERSION 5.00
Begin VB.UserControl ProgressBar 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   372
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   0
      Width           =   3852
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Bcwan Progress Bar Control
' Version 1.0
' Copyright (c) 2000 BIDA Chikh (BCwan)
' http://www.bcwansoft.com/
' bcwan@hotmail.com

Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal hpal As Long, colorref As Long) As Long

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Const m_def_Orientation = 0
Const m_def_Value = 50
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Text = "Progress"

Dim m_Orientation As Variant
Dim m_Value As Single
Dim m_Min As Single
Dim m_Max As Single
Dim m_Text As String

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Enum MyBorderStyle
  None
  Fixed
End Enum

Public Enum OrientationType
  zOrientationHorizontal
  zOrientationVertical
End Enum

Private Sub Picture1_Paint()
  Dim hBrush As Long, hBrushOld As Long
  Dim RC As RECT, sString As String
  Dim nWidth As Long, nHeight As Long
  Dim nBackColor As Long, nForeColor As Long
    
  Dim m_Percent As Single
    
  OleTranslateColor BackColor, 0, nBackColor
  OleTranslateColor ForeColor, 0, nForeColor
    
  m_Percent = (m_Value - m_Min) / (m_Max - m_Min)
    
  sString = m_Text & " " & Format(m_Percent, "#0.00%")
  
  nWidth = Picture1.TextWidth(sString)
  nHeight = Picture1.TextHeight(sString)
    
  Picture1.ForeColor = nBackColor
  Picture1.FillColor = nForeColor
  RC.left = 0
  If m_Orientation = zOrientationHorizontal Then
    RC.right = Picture1.ScaleWidth * m_Percent
    RC.top = 0
  Else
    RC.right = Picture1.ScaleWidth
    RC.top = Picture1.ScaleHeight * (1 - m_Percent)
  End If
  RC.bottom = Picture1.ScaleHeight

  SetBkColor Picture1.hdc, nForeColor
  ExtTextOut Picture1.hdc, (Picture1.ScaleWidth - nWidth) / 2, _
                           (Picture1.ScaleHeight - nHeight) / 2, _
                           4 Or 2, RC, _
                           sString, Len(sString), ByVal 0&
            
  If m_Orientation = zOrientationHorizontal Then
    RC.left = RC.right
    RC.right = Picture1.ScaleWidth
  Else
    RC.bottom = RC.top
    RC.top = 0
  End If
  Picture1.ForeColor = nForeColor
  Picture1.FillColor = nBackColor
    
  SetBkColor Picture1.hdc, nBackColor
  ExtTextOut Picture1.hdc, (Picture1.ScaleWidth - nWidth) / 2, _
                           (Picture1.ScaleHeight - nHeight) / 2, _
                           4 Or 2, RC, _
                           sString, Len(sString), ByVal 0&
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = " Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_BackColor As OLE_COLOR)
  UserControl.BackColor() = new_BackColor
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = " Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
  Refresh
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal new_font As Font)
  Set Picture1.Font = new_font
  PropertyChanged "Font"
  Refresh
End Property

Public Property Get Value() As Single
Attribute Value.VB_Description = "Returns/sets a number that specifies the value of the ProgressBar control."
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
  m_Value = New_Value
  PropertyChanged "Value"
  Refresh
End Property

Public Property Get Text() As String
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  PropertyChanged "Text"
  Refresh
End Property

Public Property Get Min() As Single
Attribute Min.VB_Description = "Returns or sets the ProgressBar control's minimum value."
  Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Single)
  If New_Min > Max Then Err.Raise 380
  m_Min = New_Min
  PropertyChanged "Min"
  Refresh
End Property

Public Property Get Max() As Single
Attribute Max.VB_Description = "Returns or sets the ProgressBar control's maximum value."
  Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Single)
  If m_Min > New_Max Then Err.Raise 380
  m_Max = New_Max
  PropertyChanged "Max"
  Refresh
End Property

Private Sub UserControl_InitProperties()
  m_Value = m_def_Value
  m_Min = m_def_Min
  m_Max = m_def_Max
  m_Text = m_def_Text
  m_Orientation = m_def_Orientation
  UserControl.ForeColor = vbBlue
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", vbBlue)
  Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  m_Min = PropBag.ReadProperty("Min", m_def_Min)
  m_Max = PropBag.ReadProperty("Max", m_def_Max)
  m_Text = PropBag.ReadProperty("Text", m_def_Text)
  m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
  Picture1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  Picture1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
  Picture1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
End Sub

Private Sub UserControl_Resize()
  Picture1.Width = ScaleWidth
  Picture1.Height = ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, vbBlue)
  Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
  Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
  Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
  Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
  Call PropBag.WriteProperty("BorderStyle", Picture1.BorderStyle, 1)
  Call PropBag.WriteProperty("OLEDragMode", Picture1.OLEDragMode, 0)
  Call PropBag.WriteProperty("OLEDropMode", Picture1.OLEDropMode, 0)
End Sub

Public Property Get Orientation() As OrientationType
Attribute Orientation.VB_Description = "Returns or sets a value that determines the orientation (horizontal or vertical) of the object."
  Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationType)
  m_Orientation = New_Orientation
  Dim temp
    
  temp = Height
  Height = Width
  Width = temp
  PropertyChanged "Orientation"
End Property

Public Property Get BorderStyle() As MyBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = Picture1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MyBorderStyle)
  Picture1.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  Picture1.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
