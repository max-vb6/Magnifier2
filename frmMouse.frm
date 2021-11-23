VERSION 5.00
Begin VB.Form frmMouse 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTxt 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   600
      Width           =   1215
      Begin VB.Label lblTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Txt"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   140
         Width           =   1260
      End
      Begin VB.Shape shpTxt 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Timer tmrMouse 
      Interval        =   10
      Left            =   1680
      Top             =   1320
   End
End
Attribute VB_Name = "frmMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mp1 As POINTAPI, mp2 As POINTAPI

Private Sub Form_Load()
Me.Move 0, 0, Screen.Width, Screen.Height
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
Dim Ret As Long
Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
Ret = Ret Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
SetLayeredWindowAttributes Me.hWnd, vbWhite, 230, LWA_ALPHA Or LWA_COLORKEY
Me.ForeColor = ThmClr.TabOnAct
GetCursorPos mp1
frmMain.SetFocus
End Sub

Private Sub tmrMouse_Timer()
GetCursorPos mp2
picTxt.Move mp2.X * Screen.TwipsPerPixelX, mp2.Y * Screen.TwipsPerPixelY + 360
Line (mp1.X * Screen.TwipsPerPixelX, mp1.Y * Screen.TwipsPerPixelY)-(mp2.X * Screen.TwipsPerPixelX, mp2.Y * Screen.TwipsPerPixelY)
mp1.X = mp2.X
mp1.Y = mp2.Y
End Sub

Sub GetMouse(SS As String)
Select Case SS
Case "左"
lblTxt.Caption = "网页后退"
Case "右"
lblTxt.Caption = "网页前进"
Case "下左"
lblTxt.Caption = "关闭标签"
Case "下右"
lblTxt.Caption = "新建标签"
Case "下上", "上下"
lblTxt.Caption = "刷新"
Case Else
lblTxt.Caption = "无动作"
End Select
End Sub

