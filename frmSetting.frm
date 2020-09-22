VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings..."
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "500"
      Top             =   795
      Width           =   735
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   450
      Width           =   240
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Height          =   240
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   450
      Width           =   240
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   240
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   97
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "milisecond(s)"
      Height          =   195
      Left            =   3000
      TabIndex        =   10
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Information update interval"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1980
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Text color"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Free RAM color"
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPU Usage color"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Background color"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
dlgColor.Color = FRmMain.picCPULoad.BackColor
dlgColor.ShowColor
If Err Then Exit Sub
Command1.BackColor = dlgColor.Color
FRmMain.picCPULoad.BackColor = dlgColor.Color
SaveSetting App.Title, "Setting", "Background Color", dlgColor.Color
End Sub

Private Sub Command2_Click()
On Error Resume Next
dlgColor.Color = FRmMain.CPUUsageColor
dlgColor.ShowColor
If Err Then Exit Sub
Command2.BackColor = dlgColor.Color
FRmMain.CPUUsageColor = dlgColor.Color
SaveSetting App.Title, "Setting", "CPU Usage Color", dlgColor.Color
End Sub

Private Sub Command3_Click()
On Error Resume Next
dlgColor.Color = FRmMain.FreeRAMColor
dlgColor.ShowColor
If Err Then Exit Sub
Command3.BackColor = dlgColor.Color
FRmMain.FreeRAMColor = dlgColor.Color
SaveSetting App.Title, "Setting", "Free RAM Color", dlgColor.Color
End Sub

Private Sub Command4_Click()
On Error Resume Next
dlgColor.Color = FRmMain.lblData.ForeColor
dlgColor.ShowColor
If Err Then Exit Sub
Command4.BackColor = dlgColor.Color
FRmMain.lblData.ForeColor = dlgColor.Color
SaveSetting App.Title, "Setting", "Text Color", dlgColor.Color
End Sub

Private Sub Form_Load()
txtInterval = FRmMain.tmrUpdate.Interval
End Sub

Private Sub Form_Unload(Cancel As Integer)
FRmMain.tmrUpdate.Interval = Val(txtInterval)
SaveSetting App.Title, "Setting", "Update Interval", txtInterval
End Sub
