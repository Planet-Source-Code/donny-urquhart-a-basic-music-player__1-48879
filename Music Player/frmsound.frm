VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "MSWEBDVD.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsound 
   BorderStyle     =   0  'None
   Caption         =   "Sound Settings"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmsound.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMute 
      Caption         =   "Mute ON"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin MSWEBDVDLibCtl.MSWebDVD snd1 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
      _cx             =   661
      _cy             =   873
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   0
      WindowlessActivation=   0   'False
   End
   Begin MSComctlLib.Slider volumecontrol 
      Height          =   2895
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   5106
      _Version        =   393216
      BorderStyle     =   1
      MousePointer    =   1
      Orientation     =   1
      LargeChange     =   50
      SmallChange     =   0
      Min             =   -10000
      Max             =   10000
      TickStyle       =   2
      TickFrequency   =   1000
   End
   Begin MSComctlLib.Slider speedcontrol 
      Height          =   2895
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   5106
      _Version        =   393216
      BorderStyle     =   1
      Enabled         =   0   'False
      Orientation     =   1
      LargeChange     =   1
      Min             =   -5000
      Max             =   5000
      SelStart        =   5000
      TickStyle       =   2
      TickFrequency   =   500
      Value           =   5000
   End
   Begin MSComctlLib.Slider balancecontrol 
      Height          =   2895
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   5106
      _Version        =   393216
      BorderStyle     =   1
      Enabled         =   0   'False
      Orientation     =   1
      LargeChange     =   1
      Min             =   -5000
      Max             =   5000
      SelStart        =   5000
      TickStyle       =   2
      TickFrequency   =   500
      Value           =   5000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblbalance 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblspeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblvolume 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmsound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMute_Click()
If snd1.Mute = True Then
cmdMute.Caption = "Mute ON"
snd1.Mute = False
ElseIf snd1.Mute = False Then
cmdMute.Caption = "Mute OFF"
snd1.Mute = True
End If
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub speedcontrol_Scroll()
Dim speed
speed = speedcontrol.Value / 100
media.frmmain.MMC.FileName
audio.Balance = speed
End Sub

Private Sub volumecontrol_Scroll()
On Error Resume Next
Dim vol1
vol1 = volumecontrol.Value - 2500
snd1.Volume = vol1
End Sub
