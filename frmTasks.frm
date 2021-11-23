VERSION 5.00
Begin VB.Form frmTasks 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "任务管理器"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "frmTasks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picNone 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2160
      ScaleHeight     =   495
      ScaleWidth      =   2895
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "没有正在运行的应用"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2835
      End
   End
   Begin VB.Timer tmrRfsh 
      Interval        =   500
      Left            =   0
      Top             =   5040
   End
   Begin Magnifier.ucCmdBtn btnRfsh 
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "刷新"
   End
   Begin Magnifier.ucCmdBtn btnEndTsk 
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "结束任务"
   End
   Begin VB.VScrollBar sroApp 
      Height          =   4215
      Left            =   6960
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTasks 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   7215
      TabIndex        =   3
      Top             =   480
      Width           =   7215
      Begin VB.PictureBox picLst 
         BackColor       =   &H00EBEBEB&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2535
         ScaleWidth      =   7215
         TabIndex        =   4
         Top             =   0
         Width           =   7215
         Begin VB.Label lblTasks 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AppTask"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblSel 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   0
            TabIndex        =   9
            Tag             =   "1"
            Top             =   0
            Width           =   7215
         End
      End
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   0
      Left            =   6720
      Picture         =   "frmTasks.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Index           =   1
      Left            =   6240
      Picture         =   "frmTasks.frx":007E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magnifier 任务管理器"
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
      TabIndex        =   0
      Top             =   80
      Width           =   2100
   End
   Begin VB.Label lblCtrl 
      BackColor       =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCtrl 
      BackColor       =   &H000000C0&
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEndTsk_Click()
On Error Resume Next
frmSHide.SC(CInt(lblSel.Tag)).Reset
Unload frmSHide.SC(CInt(lblSel.Tag))
RfshLst
End Sub

Private Sub btnRfsh_Click()
RfshLst
End Sub

Private Sub Form_Load()
RfshLst
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
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

Private Sub imgCtrl_Click(Index As Integer)
Select Case Index
Case 0
Unload Me
Case 1
Me.WindowState = 1
End Select
End Sub

Sub RfshLst()
On Error GoTo rfErr
Dim i As Long, sName As String, lT As Long
For i = 0 To lblTasks.UBound
If i > 0 Then Unload lblTasks(i)
Next i
sroApp.Visible = False
picNone.Visible = False
lblSel.Visible = True
If frmSHide.SC.Count = 1 Then
lblSel.Visible = False
picNone.Visible = True
Exit Sub
End If
For i = 1 To frmSHide.SC.Count - 1
sName = frmSHide.SC(i).Tag
Load lblTasks(lblTasks.UBound + 1)
With lblTasks(lblTasks.UBound)
.Move lblTasks(0).Left, (lblTasks.UBound - 1) * lblSel.Height + 120
.Caption = sName
.Tag = i
.Visible = True
.ZOrder 0
End With
Next i
picLst.Height = (frmSHide.SC.Count - 1) * lblSel.Height
If picLst.Height > picTasks.ScaleHeight Then
With sroApp
.Visible = True
.Max = picTasks.ScaleHeight - picLst.Height
.LargeChange = lblSel.Height * 2
.SmallChange = lblSel.Height
End With
End If
Exit Sub
rfErr:
i = i + 1
Resume
End Sub

Private Sub lblTasks_Click(Index As Integer)
lblSel.Top = lblTasks(Index).Top - 120
lblSel.Tag = Index
End Sub

Private Sub sroApp_Change()
picLst.Top = sroApp.Value
End Sub

Private Sub sroApp_Scroll()
sroApp_Change
End Sub

Private Sub tmrRfsh_Timer()
If frmTasks.Visible Then RfshLst
End Sub
