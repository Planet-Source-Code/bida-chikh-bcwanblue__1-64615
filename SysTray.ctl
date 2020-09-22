VERSION 5.00
Begin VB.UserControl SysTray 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   300
   ScaleWidth      =   255
End
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

Private Const WM_MOUSEMOVE As Long = &H200

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private VBGTray As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Default Property Values:
Const m_def_AutohideForm = True
Const m_def_Enabled = False
Const m_def_ToolTip = ""

'Property Variables:
Dim m_AutohideForm As Boolean
Dim m_Icon As Picture
Dim m_Enabled As Boolean
Dim m_ToolTip As String
Dim m_Minimized As Boolean
Dim TrayI As NOTIFYICONDATA

'Event Declarations:
Event LeftButtonDown()
Event LeftButtonUp()
Event LeftButtonDblClick()
Event RightButtonDown()
Event RightButtonUp()
Event RightButtonDblClick()
Event MiddleButtonUp()
Event MiddleButtonDown()
Event MiddleButtonDblClick()

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim uMsg As Long
    If m_Minimized Then
        uMsg = X / Screen.TwipsPerPixelX
        Select Case uMsg
        Case WM_LBUTTONDBLCLK: RaiseEvent LeftButtonDblClick
        Case WM_LBUTTONDOWN: RaiseEvent LeftButtonDown
        Case WM_LBUTTONUP: RaiseEvent LeftButtonUp
        Case WM_MBUTTONDBLCLK: RaiseEvent MiddleButtonDblClick
        Case WM_MBUTTONDOWN: RaiseEvent MiddleButtonDown
        Case WM_MBUTTONUP: RaiseEvent MiddleButtonUp
        Case WM_RBUTTONDBLCLK: RaiseEvent RightButtonDblClick
        Case WM_RBUTTONDOWN: RaiseEvent RightButtonDown
        Case WM_RBUTTONUP: RaiseEvent RightButtonUp
        End Select
    End If
End Sub

Private Sub UserControl_Resize()
    If Width <> 420 Then Width = 420
    If Height <> 420 Then Height = 420
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error GoTo ExitSub
    m_Enabled = New_Enabled
    If Ambient.UserMode Then
        If m_Enabled Then
            Minimize
        Else
            Restore
        End If
    End If
    PropertyChanged "Enabled"
ExitSub:
End Property

Public Property Get ToolTip() As String
    ToolTip = m_ToolTip
End Property

Public Property Let ToolTip(ByVal New_ToolTip As String)
    m_ToolTip = New_ToolTip
    SetToolTip
    PropertyChanged "ToolTip"
End Property

Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_ToolTip = m_def_ToolTip
    Set m_Icon = LoadPicture("")
    m_AutohideForm = m_def_AutohideForm
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ToolTip = PropBag.ReadProperty("ToolTip", m_def_ToolTip)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_AutohideForm = PropBag.ReadProperty("AutohideForm", m_def_AutohideForm)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ToolTip", m_ToolTip, m_def_ToolTip)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("AutohideForm", m_AutohideForm, m_def_AutohideForm)
End Sub

Private Sub Minimize()
    If Not m_Minimized Then
        With TrayI
            .uId = vbNull
            .cbSize = Len(TrayI)
            .hWnd = UserControl.hWnd
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .hIcon = IIf(m_Icon Is Nothing, 0, m_Icon)
            .uCallbackMessage = WM_MOUSEMOVE
            .szTip = m_ToolTip & vbNullChar
            Call Shell_NotifyIcon(NIM_ADD, TrayI)
        End With
        m_Minimized = True
        m_Enabled = True
        If Ambient.UserMode Then If m_AutohideForm Then UserControl.Parent.Hide
    End If
End Sub

Private Sub SetToolTip()
    With TrayI
        .cbSize = Len(VBGTray)
        .hWnd = UserControl.hWnd
        .szTip = m_ToolTip
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, TrayI)
End Sub

Private Sub Restore()
    If m_Minimized Then
        With TrayI
            .cbSize = Len(VBGTray)
            .hWnd = UserControl.hWnd
            .uId = vbNull
        End With
        Call Shell_NotifyIcon(NIM_DELETE, TrayI)
        m_Minimized = False
        m_Enabled = False
    End If
End Sub

Private Sub SetIcon()
    With TrayI
        .cbSize = Len(VBGTray)
        .hWnd = UserControl.hWnd
        .hIcon = m_Icon
    End With
    Call Shell_NotifyIcon(NIM_MODIFY, TrayI)
End Sub

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
End Property

Public Property Get AutohideForm() As Boolean
    AutohideForm = m_AutohideForm
End Property

Public Property Let AutohideForm(ByVal New_AutohideForm As Boolean)
    m_AutohideForm = New_AutohideForm
    PropertyChanged "AutohideForm"
End Property
