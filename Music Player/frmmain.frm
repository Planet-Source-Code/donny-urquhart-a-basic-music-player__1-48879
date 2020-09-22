VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "MediaZonian"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0442
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load..."
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdVolume 
      Caption         =   "Volume"
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton loop1 
      Caption         =   "Loop ON"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstmedia1 
      Height          =   2835
      Left            =   7320
      TabIndex        =   5
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   5001
      View            =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "pics"
      SmallIcons      =   "pics"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList pics 
      Left            =   4320
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":BAD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MCI.MMControl MMC 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   6600
      Top             =   5040
   End
   Begin MSComctlLib.ProgressBar tracksofar 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog songdlg 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox trackpic 
      Height          =   2835
      Left            =   120
      Picture         =   "frmmain.frx":BF2C
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   1
      Top             =   720
      Width           =   7140
   End
   Begin VB.Label lblplaylist 
      BackStyle       =   0  'Transparent
      Caption         =   "PlayList:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Image aboutover 
      Height          =   255
      Left            =   6120
      Picture         =   "frmmain.frx":56392
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image aboutmain 
      Height          =   255
      Left            =   6360
      Picture         =   "frmmain.frx":5662F
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image about 
      Height          =   255
      Left            =   7800
      Picture         =   "frmmain.frx":567FC
      Top             =   105
      Width           =   255
   End
   Begin VB.Label trackname 
      BackStyle       =   0  'Transparent
      Caption         =   "TRACK: "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7095
   End
   Begin VB.Image minimizeover 
      Height          =   255
      Left            =   6600
      Picture         =   "frmmain.frx":569C9
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image minimizemain 
      Height          =   255
      Left            =   6840
      Picture         =   "frmmain.frx":56C50
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image closeover 
      Height          =   255
      Left            =   7080
      Picture         =   "frmmain.frx":56DFF
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image closemain 
      Height          =   255
      Left            =   7320
      Picture         =   "frmmain.frx":5709E
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image minimize 
      Height          =   255
      Left            =   8160
      Picture         =   "frmmain.frx":5733C
      Top             =   105
      Width           =   255
   End
   Begin VB.Image closepic 
      Height          =   255
      Left            =   8520
      Picture         =   "frmmain.frx":574EB
      Top             =   105
      Width           =   255
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'You may use this code freely but the pictures
'are copyrighted. Have Fun :)
Dim looper As Boolean

Private Sub about_Click()
frmabout.Show
End Sub

Private Sub about_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
about.Picture = aboutover.Picture
End Sub

Private Sub closepic_Click()
End
End Sub

Private Sub closepic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closepic.Picture = closeover.Picture
End Sub

Private Sub cmdload_Click()
On Error Resume Next
songdlg.Filter = "MP3|*.mp3|ASF|*.asf|ASX|*.asx|WAV|*.wav|AVI|*.avi|"
songdlg.ShowOpen

  MMC.FileName = songdlg.FileName
  MMC.Command = "close"
  MMC.Command = "open"
  
  tracksofar.Max = MMC.Length
  tracksofar.Min = 0
  
  trackname.Caption = "TRACK: " & songdlg.FileTitle
  lstmedia1.ListItems.Add , , songdlg.FileTitle, , 1
End Sub

Private Sub cmdpause_Click()
If cmdPause.Caption = "Pause" Then
cmdPause.Caption = "Resume"
ElseIf cmdPause.Caption = "Resume" Then
cmdPause.Caption = "Pause"
End If
    MMC.Command = "pause"
End Sub

Private Sub cmdplay_Click()
      MMC.Command = "seek"
      MMC.Command = "play"
End Sub

Private Sub cmdstop_Click()
    MMC.Command = "stop"
    cmdPause.Caption = "Pause"
End Sub

Private Sub cmdVolume_Click()
frmsound.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
App.TaskVisible = True
frmsound.snd1.Volume = 25000
frmsound.volumecontrol.Value = 50
MMC.Shareable = True
looper = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
closepic.Picture = closemain.Picture
minimize.Picture = minimizemain.Picture
about.Picture = aboutmain.Picture
End Sub

Private Sub loop1_Click()
If loop1.Caption = "Loop ON" Then
looper = False
loop1.Caption = "Loop OFF"
ElseIf loop1.Caption = "Loop OFF" Then
loop1.Caption = "Loop ON"
looper = True
End If
End Sub

Private Sub lstmedia1_DblClick()
MMC.FileName = lstmedia1.SelectedItem
trackname.Caption = "TRACK: " & lstmedia1.SelectedItem
MMC.Command = "close"
MMC.Command = "open"
MMC.Command = "play"
End Sub

Private Sub lstmedia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
about.Picture = aboutmain.Picture
closepic.Picture = closemain.Picture
minimize.Picture = minimizemain.Picture
End Sub

Private Sub minimize_Click()
frmmain.WindowState = vbMinimized
End Sub

Private Sub minimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
minimize.Picture = minimizeover.Picture
End Sub

Private Sub Timer1_Timer()
tracksofar.Value = MMC.Position
If tracksofar.Value = MMC.TrackLength Then
    If looper = True Then
        MMC.Command = "close"
        MMC.Command = "open"
        MMC.Command = "play"
    End If
End If
End Sub

Private Sub Vol_Change()
Volcaption.Caption = vol.Value
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub trackname_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

