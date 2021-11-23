VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MagDown"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Dwn"
   ScaleHeight     =   6150
   ScaleWidth      =   4815
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picDDE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin MagDown.ucDwnLst DW 
      Height          =   5670
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   10001
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   3840
      Picture         =   "frmMain.frx":4781A
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   4320
      Picture         =   "frmMain.frx":47879
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magnifier 下载管理器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   80
      Width           =   2100
   End
   Begin VB.Label lblCtrl 
      BackColor       =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCtrl 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
On Error Resume Next
Me.Show
Me.ZOrder 0
Me.SetFocus
If LCase(CmdStr) <> "show" Then
Dim Strs() As String
Strs = Split(CmdStr, "@@")
DW.AddNewDwn Strs(0), Strs(1)
End If
Cancel = False
End Sub

Private Sub Form_Load()
On Error GoTo LoadErr:
If Command = "" Then End
If App.PrevInstance Then
Me.LinkTopic = ""
Me.LinkMode = 0
LinkAndSendMessage Command
End
End If
SetClassLong Me.hwnd, GCL_STYLE, GetClassLong(Me.hwnd, GCL_STYLE) Or CS_DROPSHADOW  'Form has a shadow
Me.Show
If LCase(Command) <> "show" Then
Dim Strs() As String
Strs = Split(Command, "@@")
DW.AddNewDwn Strs(0), Strs(1)
End If
Exit Sub
LoadErr:
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub LinkAndSendMessage(ByVal Msg As String)
On Error GoTo errH
Dim t As Long
With picDDE
.LinkMode = 0
.LinkTopic = "MagDown|Dwn"
.LinkMode = 2
.LinkExecute Msg

t = .LinkTimeout
.LinkTimeout = 1
.LinkMode = 0
.LinkTimeout = t
End With
Exit Sub
errH:
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("关闭下载管理器会导致所有正在下载的任务停止" & vbCrLf & "是否继续？", 32 + vbYesNo, "关闭下载管理器") = vbNo Then Cancel = 1
End Sub

Private Sub imgCtrl_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
Case 1
Me.WindowState = 1
End Select
End Sub

Private Sub imgCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCtrl(Index).Visible = True
End Sub

Private Sub imgCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCtrl(Index).Visible = False
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub
