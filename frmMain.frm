VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bcwan Folder Notifier and BlueTooth Sender"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   6930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin BcwanBlueTooth.SysTray SysTray1 
      Left            =   7080
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      Icon            =   "frmMain.frx":0ECA
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   720
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   2415
      Index           =   0
      Left            =   120
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      FrameColor      =   192
      BackColor       =   13160660
      FillColor       =   14215660
      TextBoxColor    =   192
      TxtBoxShadow    =   1
      Style           =   5
      Caption         =   "       Folder"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin BcwanBlueTooth.CButton SelectFolderCommand 
         Height          =   285
         Left            =   6240
         TabIndex        =   16
         ToolTipText     =   "Select the Folder to Monitor."
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   16777152
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16744576
         MouseIcon       =   "frmMain.frx":1DA4
         MousePointer    =   99
      End
      Begin VB.CheckBox IncludeSubDirCheck 
         Caption         =   "Include Sub Folders"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmMain.frx":20BE
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1245
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "List of all changes"
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox FolderToMonitorText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Text            =   "c:"
         Top             =   480
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":23C8
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         Caption         =   "Folder"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   525
         Width           =   615
      End
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   120
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   2535
      Index           =   1
      Left            =   120
      Top             =   2520
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      FrameColor      =   192
      BackColor       =   12632319
      FillColor       =   14215660
      TextBoxColor    =   192
      TxtBoxShadow    =   1
      Style           =   5
      Caption         =   "       BlueTooth Device"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      GradientHeaderStyle=   1
      Begin VB.ComboBox SendToBlueToothCombo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox SenderFileText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox DeviceAdressText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "00:00:00:00:00:00"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox DeviceNameText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "MyDevice0x201"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox DeviceClassText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "00000000"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tip : These Informations can be grabbed from the Shortcut ""SendTo"" to your device "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   6495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "The Device Class."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "The Name identiying the device."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "It's a 6 couples of Hexa digits like this. (*#2820#) ."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":28CA
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sender File :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Device Adress :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Device Name :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Device Class :"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   1455
      Index           =   2
      Left            =   120
      Top             =   5160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2566
      FrameColor      =   192
      BackColor       =   8438015
      FillColor       =   14215660
      TextBoxColor    =   192
      TxtBoxShadow    =   1
      Style           =   5
      Caption         =   "       Notification"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      Begin VB.OptionButton NotifyOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Caption         =   "Notify Every"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMain.frx":2D8F
         MousePointer    =   99  'Custom
         TabIndex        =   12
         ToolTipText     =   "Set the laps of time to Notify."
         Top             =   840
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton NotifyOption 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8E9EC&
         Caption         =   "Notify Every Modification (Not Recommended)"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMain.frx":3099
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "Notification Every Chane (Not Recommended)"
         Top             =   480
         Width           =   3615
      End
      Begin BcwanBlueTooth.ucUpDownBox MinutesCounter 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Min             =   1
         Max             =   60
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tip : It's not recommended that to set to notify every change (Many can raise in a second)."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   2
         Left            =   120
         Picture         =   "frmMain.frx":33A3
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   870
         Width           =   615
      End
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   1845
      Index           =   3
      Left            =   120
      Top             =   6720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3254
      FrameColor      =   192
      BackColor       =   13160660
      FillColor       =   14215660
      TextBoxColor    =   192
      TxtBoxShadow    =   1
      Style           =   5
      Caption         =   "       Actions (Send file instantly if)"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1125
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":3904
         Left            =   1440
         List            =   "frmMain.frx":3906
         MouseIcon       =   "frmMain.frx":3908
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   21
         Top             =   795
         Width           =   1215
      End
      Begin BcwanBlueTooth.CButton AddExtCommand 
         Height          =   285
         Left            =   5400
         TabIndex        =   22
         ToolTipText     =   "Add Entry to Actions List"
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         BackColor       =   16777152
         Caption         =   "Add to List"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16744576
         MouseIcon       =   "frmMain.frx":3C12
         MousePointer    =   99
      End
      Begin BcwanBlueTooth.CButton DeleteExtCommand 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Remove Selected Item from the actions List"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         BackColor       =   16777152
         Caption         =   "Remove"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16744576
         MouseIcon       =   "frmMain.frx":3F2C
         MousePointer    =   99
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Any File added with the Type (Extension) Listed Above is sent  instantly."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   6495
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   3
         Left            =   120
         Picture         =   "frmMain.frx":4246
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remove File Type from the List"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Add File Type (Extension)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   1005
      Index           =   5
      Left            =   120
      Top             =   10200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1773
      FrameColor      =   192
      BackColor       =   13160660
      FillColor       =   12640511
      TextBoxColor    =   192
      Style           =   5
      Caption         =   "       About"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      GradientHeaderStyle=   1
      Begin VB.Image EmailMe 
         Height          =   480
         Index           =   7
         Left            =   1920
         MouseIcon       =   "frmMain.frx":476C
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":4A76
         ToolTipText     =   "Email me : bcwan@hotmail.com "
         Top             =   360
         Width           =   480
      End
      Begin VB.Image WebSite 
         Height          =   480
         Index           =   6
         Left            =   4440
         MouseIcon       =   "frmMain.frx":5129
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":5433
         ToolTipText     =   "Visit my WebSite : http://www.bcwansoft.com"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   5
         Left            =   120
         Picture         =   "frmMain.frx":5AEF
         Top             =   120
         Width           =   360
      End
   End
   Begin BcwanBlueTooth.jcFrames jcFrames1 
      Height          =   1125
      Index           =   4
      Left            =   120
      Top             =   9000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1984
      FrameColor      =   192
      BackColor       =   13160660
      FillColor       =   12640511
      TextBoxColor    =   192
      TxtBoxShadow    =   1
      Style           =   5
      Caption         =   "      Informations"
      TextColor       =   16777215
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      Begin VB.Image Image1 
         Height          =   360
         Index           =   4
         Left            =   120
         Picture         =   "frmMain.frx":6015
         Top             =   120
         Width           =   360
      End
      Begin VB.Label ErrorLabel 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4815
      End
   End
   Begin BcwanBlueTooth.CButton MonitorCommand 
      Height          =   1125
      Left            =   5280
      TabIndex        =   18
      ToolTipText     =   "Start or Stop Monitoring"
      Top             =   9000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1984
      BackColor       =   16777152
      Caption         =   "Start Monitoring"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMain.frx":6504
      MaskColor       =   0
      PicCapAlign     =   4
      BorderStyle     =   1
      BorderColor     =   16744576
      MouseIcon       =   "frmMain.frx":73DE
      MousePointer    =   99
   End
   Begin BcwanBlueTooth.ProgressBar BlueProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   8640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      BackColor       =   12648447
      ForeColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   0
      Max             =   30
      BorderStyle     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Interval As Integer
Dim ThreadHandle As Long
Dim AppName As String

Private Sub AddExtCommand_Click()
  List2.AddItem UCase(Text1.Text)
  Text1.Text = ""
End Sub

Private Sub DeleteExtCommand_Click()
  List2.RemoveItem List2.ListIndex
  Label6.Visible = False
  DeleteExtCommand.Visible = False
End Sub

Private Sub EmailMe_Click(Index As Integer)
  ShellExecuteA Me.hwnd, "open", "mailto: bcwan@hotmail.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  SetBlueSettings
End Sub

Private Sub SetBlueSettings()
  Dim I As Integer
  
  SaveSetting AppName, "Folder", "FolderMonitor", FolderToMonitorText.Text
  SaveSetting AppName, "Folder", "InCludeSubs", IncludeSubDirCheck.Value

  SaveSetting AppName, "BlueTooth", "SenderExeFile", SenderFileText.Text
  SaveSetting AppName, "BlueTooth", "DeviceAdress", DeviceAdressText.Text
  SaveSetting AppName, "BlueTooth", "DeviceName", DeviceNameText.Text
  SaveSetting AppName, "BlueTooth", "DeviceClass", DeviceClassText.Text

  If NotifyOption(0).Value = True Then
    SaveSetting AppName, "Notification", "NotifyType", "0"
  Else
    SaveSetting AppName, "Notification", "NotifyType", "1"
  End If
  SaveSetting AppName, "Notification", "Interval", MinutesCounter.Value

  SaveSetting AppName, "Actions", "Much", List2.ListCount
  If List2.ListCount >= 1 Then
    For I = 0 To List2.ListCount - 1
      SaveSetting AppName, "Actions", "Action_" & Str(I), List2.List(I)
    Next
  End If
End Sub

Private Sub GetBlueSettings()
  Dim I As Integer
  Dim Mch As Long
  Dim T As String
  
  On Error Resume Next
  FolderToMonitorText.Text = GetSetting(AppName, "Folder", "FolderMonitor")
  IncludeSubDirCheck.Value = GetSetting(AppName, "Folder", "InCludeSubs")

  SenderFileText.Text = GetSetting(AppName, "BlueTooth", "SenderExeFile")
  DeviceAdressText.Text = GetSetting(AppName, "BlueTooth", "DeviceAdress")
  DeviceNameText.Text = GetSetting(AppName, "BlueTooth", "DeviceName")
  DeviceClassText.Text = GetSetting(AppName, "BlueTooth", "DeviceClass")

  
  If GetSetting(AppName, "Notification", "NotifyType") = "0" Then
    NotifyOption(0).Value = True
  Else
    NotifyOption(1).Value = True
  End If

  MinutesCounter.Value = GetSetting(AppName, "Notification", "Interval")

  Mch = GetSetting(AppName, "Actions", "Much")
  For I = 0 To Mch - 1
    T = GetSetting(AppName, "Actions", "Action_" & Str(I))
    List2.AddItem T
  Next
End Sub
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    SysTray1.Enabled = True
    SysTray1.ToolTip = "Monitoring the folder " & FolderToMonitorText.Text & " and send changes via BlueTooth"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  WatchStart = False
    SysTray1.Enabled = False
End Sub

Private Sub IncludeSubDirCheck_Click()
  WSubFolder = IncludeSubDirCheck.Value
End Sub

Private Sub FolderToMonitorText_Change()
  If FolderToMonitorText.Text <> "" Then
    IncludeSubDirCheck.Visible = True
  Else
    IncludeSubDirCheck.Visible = False
  End If
End Sub

Private Sub List2_Click()
  Label6.Visible = (List2.ListIndex >= 0)
  DeleteExtCommand.Visible = (List2.ListIndex >= 0)
End Sub

Private Sub MinutesCounter_Change()
  BlueProgressBar.Max = MinutesCounter.Value * 60
End Sub

Private Sub MonitorCommand_Click()
  Dim Dummy As Long
  Dim Changes As String
  Dim WaitNum As Long
  Dim I As Integer
  Dim Fl As String
  Dim Fm As String
  
  If UCase(MonitorCommand.Caption) = "START MONITORING" Then
    For I = 0 To 4
      jcFrames1(I).Enabled = False
    Next
    SelectFolderCommand.Visible = False
    MinutesCounter.Enabled = False
    Timer1.Enabled = False
    BlueProgressBar.Visible = True
         
    Interval = MinutesCounter.Value * 60
    BlueProgressBar.Max = Interval
    Timer.Enabled = True
    MonitorCommand.Caption = "Stop Monitoring"
    WSubFolder = IncludeSubDirCheck.Value
    WatchStart = True
    'Get Folder Handle
    FolderPath = FolderToMonitorText.Text
    If right(FolderPath, 1) <> "\" Then FolderPath = FolderPath + "\"
    DirHndl = GetDirHndl(FolderPath)
    If (DirHndl = 0) Or (DirHndl = -1) Then MsgBox "Cannot create handle": Exit Sub
    'Create thread to Watch changes
    Do
      ThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartWatch, DirHndl, 0, Dummy)
      Do
        WaitNum = WaitForSingleObject(ThreadHandle, 50)
        DoEvents
      Loop Until (WaitNum = 0) Or (WatchStart = False)
      Changes = ""
      If WaitNum = 0 Then Changes = GetChanges
      If Changes <> "" Then
        List1.AddItem Changes
        If left(Changes, 11) = "Added file-" Then
          Fl = Mid(Changes, 12, Len(Changes) - 11)
          Fm = UCase(right(Fl, 3))
          If InActionsList(Fm) Then
            SendFileToBlue Fl
            'Here wa have to wait to complete the Transfert of the file to process another one
            'More afinity
          End If
        End If
        If Timer.Enabled = False Then SendToBlue Changes
      End If
    Loop Until Not WatchStart
    'Terminate the Thread & Clear Handle
    If DirHndl <> 0 Then ClearHndl DirHndl
    If ThreadHandle <> 0 Then Call TerminateThread(ThreadHandle, ByVal 0&): ThreadHandle = 0
  Else
    Timer.Enabled = False
    WatchStart = False
    MonitorCommand.Caption = "Start Monitoring"
     For I = 0 To 4
      jcFrames1(I).Enabled = True
    Next
    SelectFolderCommand.Visible = True
    MinutesCounter.Enabled = True
    Timer1.Enabled = True
    BlueProgressBar.Visible = False
  End If
End Sub

Private Function InActionsList(Ext As String) As Boolean
  Dim I As Long
  Dim Y As Boolean
  
  Y = False
  For I = 0 To List2.ListCount - 1
    If List2.List(I) = Ext Then
      Y = True
      Exit For
    End If
  Next
  InActionsList = Y
End Function

Private Sub SelectFolderCommand_Click()
  Dim Mt As String
  
  Mt = SelectFolder(Me)
  If Mt <> "" Then
    FolderToMonitorText.Text = Mt
  End If
End Sub

Private Sub form_load()
  AppName = "BcwanBlueTooth"
  
  GetBlueSettings
  MonitorCommand.Visible = False
'  MinutesCounter.Value = 5
  BlueProgressBar.Visible = False
  AddExtCommand.Visible = False
  PopulateSendToCombo
End Sub

Private Sub NotifyOption_Click(Index As Integer)
  Select Case Index
    Case 0:
      MinutesCounter.Visible = False
      Label3.Visible = False
    Case 1:
      MinutesCounter.Visible = True
      Label3.Visible = True
  End Select
End Sub

Private Sub SendToBlueToothCombo_Click()
  ShowFileProp SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, Me
'  MsgBox "Target : " & LookPropertyLink(SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, "Target") & vbCrLf _
  & "Name : " & LookPropertyLink(SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, "Name") & vbCrLf _
  & "Icon : " & LookPropertyLink(SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, "icon") & vbCrLf _
  & "Start : " & LookPropertyLink(SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, "start") & vbCrLf _
  & "Key : " & LookPropertyLink(SendToBlueToothCombo.Tag & SendToBlueToothCombo.Text, "Key")
End Sub

Private Sub SysTray1_LeftButtonDblClick()
  Me.Show
  Me.WindowState = vbNormal
  SysTray1.Enabled = False
End Sub

Private Sub Text1_Change()
  Dim I As Integer
  Dim Vs As Boolean
  Dim T As String
  
  Vs = (Text1.Text <> "")
  If Vs Then
    T = UCase(Text1.Text)
    For I = 0 To List2.ListCount - 1
      If T = List2.List(I) Then
        Vs = False
        Exit For
      End If
    Next
  End If
  AddExtCommand.Visible = Vs
End Sub

Private Sub Timer_Timer()
  Interval = Interval - 1
  If Interval <= 1 Then
    If List1.ListCount >= 1 Then SendAllToBlue
    Interval = MinutesCounter.Value * 60
  End If
  BlueProgressBar.Value = Interval
  BlueProgressBar.Text = "Remaining time " & Interval & " Sec"
End Sub

Private Sub SendToBlue(Message As String)
  Dim Pth As String
  Dim I As Integer
  Dim F As Long
  Dim Mt As String
  
  If right(App.Path, 1) = "\" Then
    Pth = App.Path & GenerateMessageFileName
  Else
    Pth = App.Path & "\" & GenerateMessageFileName
  End If
  F = FreeFile
  Open Pth For Binary Access Write As #F
    Mt = Message
    Put #F, , Mt
  Close #F
  Pth = " " & Pth
  Shell GetBlueToothSenderPath & Pth
  List1.Clear
End Sub

Private Sub SendFileToBlue(File As String)
  Dim Pth As String
  Dim I As Integer
  Dim F As Long
  Dim Mt As String
  
  If Dir(File) <> "" Then
    File = " " & File
    Shell GetBlueToothSenderPath & File
  End If
End Sub

Private Sub SendAllToBlue()
  Dim Pth As String
  Dim I As Integer
  Dim F As Long
  Dim Mt As String
    
  Mt = GenerateMessageFileName
  If right(App.Path, 1) = "\" Then
    Pth = App.Path & Mt
  Else
    Pth = App.Path & "\" & Mt
  End If
  F = FreeFile
  Open Pth For Binary Access Write As #F
    For I = 0 To List1.ListCount - 1
      Mt = List1.List(I) & vbCrLf
      Put #F, , Mt
    Next
  Close #F
  Pth = " " & Pth
  Shell GetBlueToothSenderPath & Pth
  List1.Clear
End Sub

Private Function GetBlueToothSenderPath() As String
  GetBlueToothSenderPath = Trim(SenderFileText.Text) & _
                           " -BD_ADDR=" & DeviceAdressText.Text & _
                           " -BD_NAME=" & DeviceNameText.Text & _
                           " -DEV_CLASS=" & DeviceClassText.Text
End Function

Private Function GenerateMessageFileName() As String
  GenerateMessageFileName = "Bcwan" & "-" & right("0" & Hour(Time), 2) & "-" & right("0" & Minute(Time), 2) & "-" & right("0" & Second(Time), 2) & ".txt"
End Function

Private Function TestIntegrity() As Boolean
  TestIntegrity = False
  ErrorLabel.Caption = ""
  If FolderToMonitorText.Text = "" Then
    ErrorLabel.Caption = "You have to select a folder to monitor ..."
'    FolderToMonitorText.SetFocus
    Exit Function
  End If
  If SenderFileText.Text = "" Then
    ErrorLabel.Caption = "You have to set the Bluetooth sender executable file ..."
    SenderFileText.SetFocus
    Exit Function
  End If
  If DeviceAdressText.Text = "" Then
    ErrorLabel.Caption = "You have to set the Bluetooth Device Adress " & vbCrLf & "Tip : type *#2820# on Nokia Mobiles to get it"
    DeviceAdressText.SetFocus
    Exit Function
  End If
  If DeviceNameText.Text = "" Then
    ErrorLabel.Caption = "You have to set the Bluetooth Device Name " & vbCrLf & "Tip : include spaces or special characters (ex: 0x20 for space)"
    DeviceNameText.SetFocus
    Exit Function
  End If
  If DeviceClassText.Text = "" Then
    ErrorLabel.Caption = "You have to set the Bluetooth Device Class " & vbCrLf & "Read the Readme.txt file to get more Details"
    DeviceClassText.SetFocus
    Exit Function
  End If
 
  TestIntegrity = True
End Function

Private Sub Timer1_Timer()
  If TestIntegrity Then
    MonitorCommand.Visible = True
    jcFrames1(4).Width = 5055
  Else
    MonitorCommand.Visible = False
    jcFrames1(4).Width = 6735
  End If
End Sub

Private Sub WebSite_Click(Index As Integer)
  ShellExecuteA Me.hwnd, "open", "http://www.bcwansoft.com", "", "", SW_SHOW Or SW_NORMAL
End Sub

Private Sub PopulateSendToCombo()
  Dim strP As String
  Dim strA As String
  
  strP = fGetSpecialFolder(CSIDL_SENDTO)
  If right(strP, 1) <> "\" Then strP = strP + "\"
  strP = strP + "Bluetooth\"
  ' strA=Dir$(strP+"\*.exe", vbDirectory) ' *.exe files only (No Folders)
  ' add vbDirectory to include folder names
  
  SendToBlueToothCombo.Clear
  SendToBlueToothCombo.Tag = strP
  strA = Dir(strP, vbDirectory) 'get first file
  If UCase(right(strA, 4)) = ".LNK" Then SendToBlueToothCombo.AddItem strA
  While strA > ""
    If strA <> "." And strA <> ".." Then
      If GetAttr(strP + strA) And vbDirectory Then
        ' a folder
      Else
        ' a file
        If UCase(right(strA, 4)) = ".LNK" Then SendToBlueToothCombo.AddItem strA
      End If
    End If
    strA = Dir ' Get Next File
  Wend
End Sub

Private Function LookPropertyLink(TheFullPath As String, Wanted As String) As String
  'Easy read for shortcut properties!
  'Tested:Windows XP
  'VB6
    
  'Important...! You should activate the (Windows Script Host Object Model) reference
  'in Menu (Project --> References)
  
  'IN: TheFullPath = provide the path and name of file with extension .lnk
  'IN: Wanted = property in shortcut you want
  
  'OUT: With select case and nothing ("") if path not found.
    
  'Call example... MsgBox "Target" & ": " & LookPropertyLink("c:\test.lnk", "target")
  Dim LinkShell As New WshShell
  Dim LinkShortCut As New WshShortcut
  Set LinkShortCut = LinkShell.CreateShortcut(TheFullPath)

  Select Case UCase(Wanted)
    Case "TARGET"
      LookPropertyLink = LinkShortCut.TargetPath
    Case "NAME"
      LookPropertyLink = LinkShortCut.FullName
    Case "ICON"
      LookPropertyLink = LinkShortCut.IconLocation
    Case "START"
      LookPropertyLink = LinkShortCut.WorkingDirectory
    Case "KEY"
      LookPropertyLink = LinkShortCut.Hotkey 'if any
    Case Else
  End Select
  Set LinkShell = Nothing
  Set LinkShortCut = Nothing
End Function

Private Function ShowFileProp(ByVal FileName As String, aForm As Form) As Long
  Dim SEI As SHELLEXECUTEINFO
  Dim r As Long

  If FileName = "" Then
    ShowFileProp = 0
    Exit Function
  End If
  With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = aForm.hwnd
    .lpVerb = "properties"
    .lpFile = FileName
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
  End With
  r = ShellExecuteEX(SEI)
  ShowFileProp = SEI.hInstApp
End Function

