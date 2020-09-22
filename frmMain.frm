VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "CPURAM"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   5.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   1800
   End
   Begin VB.PictureBox picCPULoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin ComctlLib.ImageList imgList 
         Left            =   720
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "Settings..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 200 'Replace the szTip string's length with your tip's length
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private UsedPhysicalMemory As Long
Private TotalPhysicalMemory As Long
Private AvailablePhysicalMemory As Long
Private TotalPageFile As Long
Private AvailablePageFile As Long
Private TotalVirtualMemory As Long
Private AvailableVirtualMemory As Long

Private m_oCPULoad As CPULoad
Private m_lCPUs As Long

Public CPUUsageColor As OLE_COLOR, FreeRAMColor As OLE_COLOR

Private Sub Form_Load()
    frmAbout.Show
    frmAbout.Refresh
    If App.PrevInstance = True Then
        MsgBox "CPURAM is already running!", vbInformation, "CPURAM"
        End
    End If
    
    CPUUsageColor = GetSetting(App.Title, "Setting", "CPU Usage Color", vbRed)
    FreeRAMColor = GetSetting(App.Title, "Setting", "Free RAM Color", vbGreen)

    picCPULoad.BackColor = GetSetting(App.Title, "Setting", "Background Color", vbBlack)
    lblData.ForeColor = GetSetting(App.Title, "Setting", "Text Color", vbWhite)
    
    tmrUpdate.Interval = Val(GetSetting(App.Title, "Setting", "Update Interval", 500))
    
    Set m_oCPULoad = New CPULoad
    m_lCPUs = m_oCPULoad.GetCPUCount
    
    StayOnTop Me
    
    CreateIcon
    
    tmrUpdate.Enabled = True

    Me.Hide
    Unload frmAbout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
picCPULoad.Move 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1
lblData.Move 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1

'        picCPULoad.Line (0, 0)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbBlack, B
'        picCPULoad.Line (picCPULoad.ScaleWidth - 1, 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight), vbWhite
'        picCPULoad.Line (1, picCPULoad.ScaleHeight - 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbWhite

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAbout
Unload frmSetting
DeleteIcon
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuSetting_Click()
frmSetting.Show
End Sub

Private Sub mnuShow_Click()
Me.Show
End Sub

Private Sub picCPULoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Me.Visible Then Exit Sub
    Select Case X '/ Screen.TwipsPerPixelX
    Case Is = WM_LBUTTONDOWN
        'Add the code for the left mouse click on the tray icon
        mnuShow_Click
    Case Is = WM_RBUTTONDOWN
        'Add the code for the left mouse click on the tray icon
        PopupMenu mnuSysTray
    End Select
End Sub

Private Sub tmrUpdate_Timer()
tmrUpdate.Enabled = False
    
    DoEvents

    Dim lCPULoad As Long
    Dim lCPUIndex As Long
    
    m_oCPULoad.CollectCPUData
    
    lblData.Caption = "Processor:-" & vbCrLf
    For lCPUIndex = 1 To m_lCPUs
        lCPULoad = lCPULoad + m_oCPULoad.GetCPUUsage(lCPUIndex)
        If Me.Visible Then lblData.Caption = lblData.Caption & "Processor " & Format(lCPUIndex, "@@") & ": " & Format(m_oCPULoad.GetCPUUsage(lCPUIndex), "@@@") & "%" & vbCrLf
    Next lCPUIndex
    If Me.Visible Then lblData = lblData & "------------------" & vbCrLf & "Average     : " & Format(lCPULoad, "@@@") & "%"
    
    With picCPULoad
        GetMemoryInfo
        
        .Cls
        picCPULoad.Line (1, .ScaleHeight - 2)-(.ScaleWidth / 2 - 1, .ScaleHeight + 1 - ((.ScaleHeight - 1) * (lCPULoad / m_lCPUs) / 100)), CPUUsageColor, BF
        picCPULoad.Line (.ScaleWidth / 2, .ScaleHeight - 2)-(.ScaleWidth - 2, .ScaleHeight + 1 - ((.ScaleHeight - 1) * AvailablePhysicalMemory / TotalPhysicalMemory)), FreeRAMColor, BF

        lblData.Caption = lblData & vbCrLf & vbCrLf & "Memory (RAM):-" & vbCrLf
        lblData.Caption = lblData & "Total RAM    : " & TotalPhysicalMemory / 1024 & " KB" & vbCrLf
        lblData.Caption = lblData & "Available RAM: " & Format(AvailablePhysicalMemory / 1024, String(Len(CStr(TotalPhysicalMemory / 1024)), "@")) & " KB" & vbCrLf
        
        imgList.ListImages.Remove 1
        picCPULoad.Line (0, 0)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbBlack, B
        picCPULoad.Line (picCPULoad.ScaleWidth - 1, 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight), vbWhite
        picCPULoad.Line (1, picCPULoad.ScaleHeight - 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbWhite
        imgList.ListImages.Add , , picCPULoad.Image
    End With
    
    ModifyIcon

tmrUpdate.Enabled = True
End Sub

Public Sub GetMemoryInfo()
  Dim MemStatus As MEMORYSTATUS
  MemStatus.dwLength = Len(MemStatus)
  GlobalMemoryStatus MemStatus
  UsedPhysicalMemory = MemStatus.dwMemoryLoad
  TotalPhysicalMemory = MemStatus.dwTotalPhys
  AvailablePhysicalMemory = MemStatus.dwAvailPhys
  TotalPageFile = MemStatus.dwTotalPageFile
  AvailablePageFile = MemStatus.dwAvailPageFile
  TotalVirtualMemory = MemStatus.dwTotalVirtual
  AvailableVirtualMemory = MemStatus.dwAvailVirtual
End Sub

Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub CreateIcon() 'Call this method to create the tray icon
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = picCPULoad.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = imgList.ListImages.Item(1).ExtractIcon
    Tic.szTip = lblData.Caption
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Sub ModifyIcon() 'Call this method to modify the trat icon properties
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = picCPULoad.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = imgList.ListImages.Item(1).ExtractIcon
    Tic.szTip = lblData.Caption
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub

Sub DeleteIcon() 'Call this method to remove the tray icon
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hwnd = picCPULoad.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

