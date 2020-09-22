VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   2655
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "You may use this code freely but the pictures are copyrighted. Have Fun :)"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblemail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email me at: durquh02@hotmail.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblweb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visit: www.angelfire.com/games4/durquhart/index.html"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label lblcreateed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Created By: Donny Urquhart"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblversion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.00 Beta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblmediazone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MediaZone"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub lblcreateed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub lblemail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub lblmediazone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub lblversion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub lblweb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub
