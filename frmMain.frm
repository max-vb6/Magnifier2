VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFCB9B&
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   11955
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Brw"
   ScaleHeight     =   8265
   ScaleWidth      =   11955
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrSBM 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   4920
   End
   Begin Magnifier.ucBookmarks BM 
      Height          =   2910
      Left            =   6360
      TabIndex        =   26
      Top             =   4920
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   10001
   End
   Begin Magnifier.ucListBox lstAddr 
      Height          =   1935
      Left            =   0
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
      _extentx        =   8493
      _extenty        =   3413
   End
   Begin VB.PictureBox picAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFCB9B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   0
      Picture         =   "frmMain.frx":4781A
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   7
      ToolTipText     =   "地址"
      Top             =   480
      Width           =   375
      Begin VB.PictureBox picZoom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FEDCC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   -10
         Width           =   1935
         Begin VB.Image imgDwn 
            Height          =   285
            Left            =   1560
            Picture         =   "frmMain.frx":47DB4
            ToolTipText     =   "查看已下载的文档和程序"
            Top             =   0
            Width           =   285
         End
         Begin VB.Image imgBM 
            Height          =   285
            Left            =   1200
            Picture         =   "frmMain.frx":47E12
            ToolTipText     =   "点击将网页加入收藏, 双击打开收藏夹"
            Top             =   0
            Width           =   285
         End
         Begin VB.Label lblZmVl 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
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
            Left            =   340
            TabIndex        =   16
            Top             =   20
            Width           =   480
         End
         Begin VB.Label lblPlMi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   960
            TabIndex        =   15
            Top             =   -30
            Width           =   105
         End
         Begin VB.Label lblPlMi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   80
            TabIndex        =   14
            Top             =   -30
            Width           =   180
         End
      End
      Begin VB.Timer tmrSTxt 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -240
         Top             =   -240
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FEDCC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.Image imgWebIco 
         Height          =   285
         Index           =   0
         Left            =   480
         Picture         =   "frmMain.frx":48050
         ToolTipText     =   "网页为未经过加密的普通网页"
         Top             =   0
         Width           =   285
      End
      Begin VB.Image imgWebIco 
         Height          =   285
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":48129
         ToolTipText     =   "网页经过安全加密"
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.PictureBox picApp 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   5280
      ScaleHeight     =   4335
      ScaleWidth      =   6015
      TabIndex        =   18
      Top             =   480
      Width           =   6015
      Begin VB.VScrollBar sroApp 
         Height          =   4335
         LargeChange     =   100
         Left            =   5760
         Max             =   100
         SmallChange     =   10
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picAppCore 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBEBEB&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   6015
         TabIndex        =   21
         Top             =   0
         Width           =   6015
         Begin VB.PictureBox picAIco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   0
            Left            =   0
            Picture         =   "frmMain.frx":48202
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   22
            Top             =   0
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblAddApp 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "添加应用"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   420
            Index           =   1
            Left            =   4320
            TabIndex        =   24
            ToolTipText     =   "添加应用"
            Top             =   120
            Width           =   1260
         End
      End
      Begin VB.Timer tmrSApp 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   1800
      End
      Begin VB.Label lblAddApp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "添加应用"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   0
         Left            =   4320
         TabIndex        =   25
         ToolTipText     =   "添加应用"
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应用库里空空如也..."
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   570
         Index           =   2
         Left            =   1185
         TabIndex        =   19
         Top             =   1800
         Width           =   3825
      End
   End
   Begin VB.PictureBox picCaption 
      Align           =   1  'Align Top
      BackColor       =   &H00FFCB9B&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   11955
      Begin VB.PictureBox picTabBar 
         BackColor       =   &H00FFCB9B&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         ScaleHeight     =   495
         ScaleWidth      =   6495
         TabIndex        =   2
         Top             =   0
         Width           =   6495
         Begin Magnifier.ucTabs Tabs 
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
            _extentx        =   1085
            _extenty        =   873
         End
         Begin VB.Image imgAddWeb 
            Appearance      =   0  'Flat
            Height          =   465
            Left            =   6120
            Picture         =   "frmMain.frx":49D46
            Top             =   0
            Width           =   330
         End
      End
      Begin VB.Image imgBtn 
         Height          =   480
         Index           =   0
         Left            =   0
         ToolTipText     =   "后退(向右拖拽前进)"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgBtn 
         Height          =   480
         Index           =   1
         Left            =   480
         ToolTipText     =   "刷新"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgBtn 
         Height          =   480
         Index           =   2
         Left            =   480
         ToolTipText     =   "停止加载"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   3
         Left            =   9240
         ToolTipText     =   "应用库"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   2
         Left            =   9720
         ToolTipText     =   "最小化"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   1
         Left            =   10200
         ToolTipText     =   "调节"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgCtrl 
         Height          =   480
         Index           =   0
         Left            =   10680
         ToolTipText     =   "关闭浏览器"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image imgBtn 
         Height          =   480
         Index           =   3
         Left            =   960
         ToolTipText     =   "工具菜单"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.PictureBox picImg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEDCC0&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1200
      ScaleHeight     =   2145
      ScaleWidth      =   3255
      TabIndex        =   17
      Top             =   490
      Visible         =   0   'False
      Width           =   3285
      Begin VB.Timer tmrRfsh 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   -120
         Top             =   -120
      End
      Begin VB.Image imgWebImg 
         Height          =   2175
         Left            =   0
         Stretch         =   -1  'True
         ToolTipText     =   "缩略图"
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Timer tmrKey 
      Interval        =   100
      Left            =   11400
      Top             =   600
   End
   Begin VB.PictureBox picSta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFCB9B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   1215
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sta..."
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.PictureBox picPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEDCC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   480
      Width           =   615
      Begin VB.PictureBox picPassCore 
         Appearance      =   0  'Flat
         BackColor       =   &H00FEDCC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   6135
         TabIndex        =   9
         Top             =   120
         Width           =   6135
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            BackColor       =   &H00FDBD93&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   1680
            PasswordChar    =   "#"
            TabIndex        =   12
            Top             =   1560
            Width           =   2625
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magnifier 处于密码保护状态"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   0
            Width           =   5625
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "请输入密码来解除锁定"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   10
            Top             =   960
            Width           =   2850
         End
      End
      Begin VB.Timer tmrSPass 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -240
         Top             =   -120
      End
   End
   Begin Magnifier.ucWeb Web 
      Height          =   4335
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      _extentx        =   9128
      _extenty        =   7646
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuMc 
         Caption         =   "保存当前页面"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMc 
         Caption         =   "打印当前页面"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMc 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMc 
         Caption         =   "新建标签页"
         Index           =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuMc 
         Caption         =   "关闭所有标签页"
         Index           =   4
      End
      Begin VB.Menu mnuMc 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuMc 
         Caption         =   "添加当前页面至书签"
         Index           =   6
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuMc 
         Caption         =   "整理书签"
         Index           =   7
      End
      Begin VB.Menu mnuMc 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuMc 
         Caption         =   "SuperHide"
         Index           =   9
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMc 
         Caption         =   "密码保护"
         Index           =   10
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuMc 
         Caption         =   "启用鼠标手势"
         Index           =   11
      End
      Begin VB.Menu mnuMc 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuMc 
         Caption         =   "程序设置"
         Index           =   13
      End
      Begin VB.Menu mnuMc 
         Caption         =   "主题管理"
         Index           =   14
      End
      Begin VB.Menu mnuMc 
         Caption         =   "任务管理器"
         Index           =   15
      End
      Begin VB.Menu mnuMc 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuMc 
         Caption         =   "关于 Magnifier"
         Index           =   17
      End
      Begin VB.Menu mnuMc 
         Caption         =   "关闭浏览器"
         Index           =   18
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NowPage As Long, SLM_Txt_W As Single, WebImgIndex As Long ', TabDraged As Boolean, TabPos As Long

Private Sub BM_BookmarkAdded(bmCaption As String)
SaveBookmark Web(NowPage).LocationURL, bmCaption
tmrSBM.Enabled = True
End Sub

Private Sub BM_BookmarkOpened(bmPath As String)
tmrSBM.Enabled = True
WebGoTo LoadFile(bmPath)
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
On Error Resume Next
Me.Show
Me.ZOrder 0
Me.SetFocus
AddPage
If Left(CmdStr, 7) = "DDEOpen" Then
If right(Replace(Replace(CmdStr, """", ""), "%1", ""), 4) = ".mmb" Then
WebGoTo LoadFile(Replace(Replace(right(CmdStr, Len(CmdStr) - 7), """", ""), "%1", ""))
ElseIf right(Replace(Replace(CmdStr, """", ""), "%1", ""), 5) = ".mapx" Then
If MsgBox("您确定要安装应用 " & GetFileName(Replace(Replace(right(CmdStr, Len(CmdStr) - 7), """", ""), "%1", "")) _
& "吗？", 32 + vbYesNo, "安装应用") = vbYes Then
InstallApp Replace(Replace(right(CmdStr, Len(CmdStr) - 7), """", ""), "%1", "")
End If
Else
WebGoTo Replace(Replace(right(CmdStr, Len(CmdStr) - 7), """", ""), "%1", "")
End If
ElseIf CmdStr = "OpenNew" Then
OnOpen
End If
Cancel = False
End Sub

Private Sub Form_Load()
If Dir(MyPath & "config.ini") = "" Then MsgBox "非常抱歉，出现了一个错误" & vbCrLf & "浏览器没有发现配置文件，所以不能为你浏览", 48, "错误": End

On Error Resume Next                 'DDE Link
If App.PrevInstance Then
Me.LinkTopic = ""
Me.LinkMode = 0
If Command <> "" Then
If right(Command, 4) = ".mmb" Then
LinkAndSendMessage "DDEOpen" & LoadFile(Command)
Else
LinkAndSendMessage "DDEOpen" & Replace(Replace(Command, """", ""), "%1", "")
End If
Else
LinkAndSendMessage "OpenNew"
End If
End
End If

On Error GoTo 0                      'End Source

SetClassLong Me.hWnd, GCL_STYLE, GetClassLong(Me.hWnd, GCL_STYLE) Or CS_DROPSHADOW  'Form has a shadow

LoadTheme
picPass.Top = -Me.ScaleHeight
Me.Show
LockWindow Me.hWnd, , , , Screen.Height / Screen.TwipsPerPixelY - GetTaskbarHeight

lblZmVl.Caption = ReadCon("Zoom") & "%"
If ReadCon("Addr") = 1 Then picAddr.Width = Me.ScaleWidth: Form_Resize
If ReadCon("AppSafe") = 1 Then frmSHide.SC(0).UseSafeSubset = True
AddPage

OnOpen
ReadCommand
If ReadCon("AutoSet") = 1 Then SetBrowser (True)
If ReadCon("OpenPass") = 1 Then mnuMc_Click (10)

If ReadCon("Mouse") = 1 Then mnuMc(11).Checked = True: InstallMouseHook
End Sub

Private Sub Form_Resize()
On Error Resume Next
picPass.Move 0, picPass.Top, Me.ScaleWidth, Me.ScaleHeight - picCaption.Height
If picPass.Top <> picCaption.Height Then picPass.Top = -Me.ScaleHeight Else picPass.Top = picCaption.Height
picPassCore.Move (picPass.Width - picPassCore.Width) / 2, (picPass.Height - picPassCore.Height) / 2
'===MoveIcon===
imgBtn(0).Move 0
imgBtn(1).Move imgBtn(0).Width
imgBtn(2).Move imgBtn(0).Width
imgBtn(3).Move imgBtn(1).Left + imgBtn(I).Width
imgCtrl(0).Move Me.ScaleWidth - imgCtrl(0).Width
imgCtrl(1).Move Me.ScaleWidth - imgCtrl(0).Width - imgCtrl(1).Width
imgCtrl(2).Move Me.ScaleWidth - imgCtrl(0).Width - imgCtrl(1).Width - imgCtrl(2).Width
imgCtrl(3).Move Me.ScaleWidth - imgCtrl(0).Width - imgCtrl(1).Width - imgCtrl(2).Width - imgCtrl(3).Width
'=====End======
lstAddr.Move txtAddr.Left, picCaption.Height + picAddr.Height
picApp.Move IIf(tmrSApp.Tag <> "", Me.ScaleWidth - picApp.Width, Me.ScaleWidth), picCaption.Height
BM.Move IIf(tmrSBM.Tag <> "", Me.ScaleWidth - BM.Width, Me.ScaleWidth), picCaption.Height + picAddr.Height
DW.Move Me.ScaleWidth - DW.Width, picCaption.Height + picAddr.Height
picTabBar.Move imgBtn(3).Left + imgBtn(3).Width + 120, 0, Me.ScaleWidth - imgBtn(0).Width * 3 - imgCtrl(0).Width * 4 - 1000
imgAddWeb.Move picTabBar.ScaleWidth - imgAddWeb.Width
ResizeTabs
picSta.Move 0, Me.ScaleHeight - picSta.Height
If Me.WindowState <> 1 Then picAddr.Move 0, picCaption.Height, IIf(picAddr.Width > 375, Me.ScaleWidth, 375)
If picAddr.Width > 375 Then
picZoom.Left = picAddr.Width - picZoom.Width - 120
txtAddr.Width = picAddr.Width - 600 - picZoom.Width - imgWebIco(0).Width
End If
If ReadCon("MinPass") = 1 And Me.WindowState = 1 Then
tmrSPass.Enabled = False
mnuMc_Click (10)
End If
End Sub

Private Sub LinkAndSendMessage(ByVal Msg As String)
On Error GoTo errH
Dim t As Long
With picCaption
.LinkMode = 0
.LinkTopic = "Magnifier|Brw"
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

Private Sub OnOpen()
Select Case ReadCon("OnOpen")
Case 0
WebGoTo ReadCon("HomePage")
Case 1
WebGoTo "about:nav"
Case 2
WebGoTo ReadCon("LastPage")
End Select
End Sub

Private Sub SetOrder(ctrl As Control)
ctrl.ZOrder 0
picAddr.ZOrder 0
lstAddr.ZOrder 0
picImg.ZOrder 0
picPass.ZOrder 0
picCaption.ZOrder 0
picApp.ZOrder 0
BM.ZOrder 0
End Sub

Private Sub SetCaption(Text As String)
SetWindowText Me.hWnd, Text
End Sub

Private Sub TabReAct(Index As Integer, UnloadThis As Boolean)
On Error GoTo errH
Dim LastIndex As Long
For I = 1 To Tabs.UBound
Tabs(I).BackColor = ThmClr.TabNoAct
If UnloadThis = True Then
If I < Index Then LastIndex = I
End If
Next I
If UnloadThis = False Then
Tabs(Index).BackColor = ThmClr.TabOnAct
Else
Tabs(LastIndex).BackColor = ThmClr.TabOnAct
'SetOrder Web(Tabs(LastIndex).Tag)
SetOrder Web(Tabs(LastIndex).Index)
'NowPage = Tabs(LastIndex).Tag
NowPage = Tabs(LastIndex).Index
txtAddr.Text = Web(NowPage).LocationURL
SetCaption Tabs(NowPage).Caption
End If
Exit Sub
errH:
If Err.Number = 340 Then
I = I + 1
Resume
End If
End Sub

Private Sub ResizeTabs()
On Error GoTo errH
Dim tLeft As Long
tLeft = 0
For I = 1 To Tabs.UBound
Tabs(I).Move tLeft, 0, (picTabBar.ScaleWidth - imgAddWeb.Width) / (Tabs.Count - 1)
tLeft = tLeft + (picTabBar.ScaleWidth - imgAddWeb.Width) / (Tabs.Count - 1)
If picAddr.Width = 375 Then
Web(I).Move 0, picCaption.Height, Me.ScaleWidth, Me.ScaleHeight - picCaption.Height
Else
Web(I).Move 0, picCaption.Height + picAddr.Height, Me.ScaleWidth, Me.ScaleHeight - picCaption.Height - picAddr.Height
End If
Next I
Exit Sub
errH:
If Err.Number = 340 Then
I = I + 1
Resume
End If
End Sub

Sub AddPage()
Load Tabs(Tabs.UBound + 1)
Load Web(Web.UBound + 1)
'Tabs(Tabs.UBound).Tag = Web.UBound
ResizeTabs
TabReAct Tabs.UBound, False
Tabs(Tabs.UBound).ProColor = ThmClr.MainColor
Tabs(Tabs.UBound).Visible = True
Web(Web.UBound).Visible = True
SetOrder Web(Web.UBound)
'NowPage = Tabs(Tabs.UBound).Tag
NowPage = Web.UBound
End Sub

Sub WebGoTo(sUrl As String)
On Error Resume Next
Web(NowPage).GoURL sUrl
Web(NowPage).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ReadCon("Mouse") = 1 Then
UninstallMouseHook
End If
If ReadCon("Unload") = 1 Then
If Web.Count - 1 > 1 Then If MsgBox("您的操作会关闭所有已经打开的标签页" & vbCrLf & "您确定要继续吗？", 64 + vbYesNo, "关闭多个标签页") = vbNo Then Cancel = 1
End If
SaveCon "LastPage", Web(NowPage).LocationURL
If picAddr.Width <> 375 Then SaveCon "Addr", 1 Else SaveCon "Addr", 0
If ReadCon("AutoDelete") = 1 Then
DeleteIEHistory
DoEvents
DeleteHistory
End If
Dim frms As Form
For Each frms In VB.Forms
Unload frms
Next
End Sub

Private Sub imgBM_Click()
If tmrSBM.Tag = "" Then BM.ShowSave Web(NowPage).LocationName
tmrSBM.Enabled = True
End Sub

Private Sub imgBM_DblClick()
If tmrSBM.Tag = "" Then BM.GetBookmarks
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static omX As Long
If Button = 1 Then
If Index = 0 And X > omX + 100 Then
imgBtn(0).Tag = "1"
imgBtn(0).Picture = LoadPicture(GetThmFolder & "Rt.bmp")
End If
Else
omX = X
End If
End Sub

Private Sub imgBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Select Case Index
Case 0
If imgBtn(0).Tag = "1" Then Web(NowPage).DoCommand 1: imgBtn(0).Tag = "": imgBtn(0).Picture = LoadPicture(GetThmFolder & "Lf.bmp"): Exit Sub
Web(NowPage).DoCommand 0
Case 1
Web(NowPage).DoCommand 3
Case 2
Web(NowPage).DoCommand 2
SetOrder imgBtn(1)
Case 3
PopupMenu mnuMain, 0, imgBtn(3).Left, picCaption.Height
End Select
End Sub

Private Sub imgCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 3
tmrSBM.Tag = 1: tmrSBM.Enabled = True
LoadApps
lblAddApp(1).Caption = "添加应用"
tmrSApp.Enabled = True
Case 2
Me.WindowState = 1
Case 1
picCaption_DblClick
Case 0
Unload Me
End Select
End Sub

Private Sub imgDwn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dir(MyPath & "MagDown.exe") <> "" Then
Shell MyPath & "MagDown.exe Show", vbNormalFocus
End If
End Sub

Private Sub lblAddApp_Click(Index As Integer)
Dim cdlg As New clsCdlg
cdlg.ShowOpen Me.hWnd, "Mapx应用安装包(*.Mapx)" & Chr(0) & "*.mapx", "载入Mapx安装包"
If cdlg.FileName <> "" Then
If cdlg.FileName <> MyPath & "Apps\" & GetFileName(cdlg.FileName) & ".mapx" Then
FileCopyEx cdlg.FileName, MyPath & "Apps\" & GetFileName(cdlg.FileName) & ".mapx"
cdlg.FileName = MyPath & "Apps\" & GetFileName(cdlg.FileName) & ".mapx"
End If
InstallApp cdlg.FileName
End If
End Sub

Private Sub lblPlMi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lZom As Long
lZom = Val(Replace(lblZmVl.Caption, "%", ""))
If Index = 0 Then
lZom = lZom + 50
If lZom > 1000 Then lZom = 1000: Exit Sub
Else
lZom = lZom - 50
If lZom < 50 Then lZom = 50: Exit Sub
End If
lblZmVl.Caption = lZom & "%"
Web(NowPage).DoCommand 5, lZom
SaveCon "Zoom", CStr(lZom)
End Sub

Private Sub lblZmVl_Click()
lblZmVl.Caption = "100%"
Web(NowPage).DoCommand 5, 100
SaveCon "Zoom", 100
End Sub

Private Sub lstAddr_MouseClicked()
txtAddr_KeyUp 13, 0
End Sub

Private Sub mnuMc_Click(Index As Integer)
On Error Resume Next
With Web(NowPage)
Select Case Index
Case 0
.DoCommand 6
Case 1
.DoCommand 7
Case 3
imgAddWeb_MouseUp 1, 0, 0, 0
Case 4
Dim it As Long
For it = 0 To Tabs.UBound
Tabs_TabClose (it)
Next it
Case 6
imgBM_Click
Case 7
imgBM_Click
imgBM_DblClick
Case 9
Me.Hide
frmSHide.Show
Case 10
picPass.Top = picCaption.Height
txtPass.SetFocus
Case 11
mnuMc(11).Checked = Not mnuMc(11).Checked
SaveCon "Mouse", Abs(CInt(mnuMc(11).Checked))
If mnuMc(11).Checked Then
InstallMouseHook
Else
UninstallMouseHook
End If
Case 13
'Settings
Case 14
'Themes
Case 15
frmTasks.Show
Case 17
frmAbout.Show 1
Case 18
Unload Me
End Select
End With
End Sub

Private Sub picAddr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picAddr.Width = 375 Then
SLM_Txt_W = Me.ScaleWidth
txtAddr.SetFocus
Else
SLM_Txt_W = 375
Web(NowPage).SetFocus
End If
tmrSTxt.Enabled = True
End Sub

Private Sub picAIco_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAddApp(1).Caption = "将应用拖拽出应用库可以删除应用"
End Sub

Private Sub picAIco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static ox!, oy!
With picAIco(Index)
If Button = 1 Then
.Move .Left - ox + X, .Top - oy + Y
Else
ox = X
oy = Y
End If
End With
End Sub

Private Sub picAIco_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If picAIco(Index).Left + picAIco(Index).Width < 0 Or picAIco(Index).Left > picAppCore.Width _
Or picAIco(Index).Top + picAIco(Index).Height < 0 Or picAIco(Index).Top > picAppCore.Height Then
DeleteApp Index
LoadApps
Else
tmrSApp.Tag = 1: tmrSApp.Enabled = True
Dim erMsg As String
erMsg = RunApp(MyPath & "Apps\" & ReadApp("app" & Index))
If erMsg <> "" Then MsgBox erMsg, 48, "应用出现错误"
End If
lblAddApp(1).Caption = "添加应用"
End Sub

Private Sub picCaption_DblClick()
Me.WindowState = IIf(Me.WindowState = 0, 2, 0)
End Sub

Private Sub picCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub imgAddWeb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AddPage
OnOpen
End Sub

Private Sub picCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuMain
End Sub

Private Sub sroApp_Change()
picAppCore.Top = -sroApp.Value
lblAddApp(1).Top = sroApp.Value
End Sub

Private Sub sroApp_Scroll()
sroApp_Change
End Sub

Private Sub Tabs_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Static ox As Long
'With Tabs(Index)
'If Button = 1 Then
'If Abs(X - ox) > 5 Then
'TabDraged = True
'.Move .Left - ox + X
'TabPos = Int((.Left + ox) / ((picTabBar.ScaleWidth - imgAddWeb.Width) / (Tabs.Count - 1))) + 1
'End If
'Else
'ox = X
'End If
'End With
End Sub

Private Sub Tabs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
If Button = 1 Then
'SetOrder Tabs(Index)
'If TabDraged Then
'If TabPos >= 1 And TabPos <= Tabs.Count - 1 And TabPos <> Int((Tabs(Index).Left) / ((picTabBar.ScaleWidth - imgAddWeb.Width) / (Tabs.Count - 1))) + 1 Then
'Dim i As Long, tmpC As String
'For i = 1 To Tabs.UBound
'If Tabs(i).Tag = TabPos Then Exit For
'If i = Tabs.UBound Then GoTo ReTab
'Next i
'tmpC = Tabs(i).Caption
'Tabs(i).Caption = Tabs(Index).Caption
'Tabs(Index).Caption = tmpC
'Tabs(i).Tag = Tabs(Index).Tag
'Tabs(Index).Tag = TabPos
'End If
'ReTab:
'TabDraged = False
'ResizeTabs
'End If
picImg.Visible = False
TabReAct Index, False
'SetOrder Web(Tabs(Index).Tag)
SetOrder Web(Index)
SetOrder picPass
NowPage = Index
'Web(Tabs(Index).Tag).SetFocus
Web(Index).SetFocus
'txtAddr.Text = Web(Tabs(Index).Tag).LocationURL
txtAddr.Text = Web(Index).LocationURL
SetCaption Tabs(Index).Caption & " - Magnifier"
If Web(Index).SecureVal <> 0 Then SetOrder imgWebIco(1) Else SetOrder imgWebIco(0)
ElseIf Button = 2 Then
picImg.Left = picTabBar.Left + Tabs(Index).Left + (Tabs(Index).Width - picImg.Width) / 2
imgWebImg.Picture = Web(Index).GetWebImg
WebImgIndex = Index
picImg.Visible = True
tmrRfsh.Enabled = True
ElseIf Button = 4 Then
Tabs_TabClose Index
End If
End Sub

Private Sub Tabs_TabClose(Index As Integer)
'NowPage = Tabs(Index).Tag
NowPage = Index
If Tabs.Count <> 2 Then
TabReAct Index, True
'Unload Web(Tabs(Index).Tag)
Unload Web(Index)
Unload Tabs(Index)
ResizeTabs
Else
TabReAct Index, False
OnOpen
End If
End Sub

Private Sub tmrKey_Timer()
If GetAsyncKeyState(VBRUN.vbKeyF4) <> 0 Then mnuMc_Click 10
If GetAsyncKeyState(VBRUN.vbKeyF2) <> 0 Then mnuMc_Click 9
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyS) Then mnuMc_Click 0
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyP) Then mnuMc_Click 1
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyT) Then mnuMc_Click 3
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyW) Then Tabs_TabClose CInt(NowPage)
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyD) Then mnuMc_Click 6
End Sub

Private Sub tmrRfsh_Timer()
On Error GoTo errH
tmrRfsh.Enabled = picImg.Visible
imgWebImg.Picture = Web(WebImgIndex).GetWebImg
Exit Sub
errH:
picImg.Visible = False
End Sub

Private Sub tmrSApp_Timer()
Select Case tmrSApp.Tag
Case ""
SpeedLessMove picApp, Me.ScaleWidth - picApp.Width, picCaption.Height, picApp.Width, picApp.Height, 5, tmrSApp, 1
Case Else
SpeedLessMove picApp, Me.ScaleWidth, picCaption.Height, picApp.Width, picApp.Height, 5, tmrSApp, ""
End Select
End Sub

Private Sub tmrSBM_Timer()
Select Case tmrSBM.Tag
Case ""
SpeedLessMove BM, Me.ScaleWidth - BM.Width, picCaption.Height + picAddr.Height, BM.Width, BM.Height, 5, tmrSBM, 1
Case Else
SpeedLessMove BM, Me.ScaleWidth, picCaption.Height + picAddr.Height, BM.Width, BM.Height, 5, tmrSBM, ""
End Select
End Sub

Private Sub tmrSPass_Timer()
SpeedLessMove picPass, 0, -Me.ScaleHeight, picPass.Width, picPass.Height, 10, tmrSPass
End Sub

Private Sub tmrSTxt_Timer()
On Error Resume Next
SpeedLessMove picAddr, picAddr.Left, picAddr.Top, SLM_Txt_W, picAddr.Height, 5, tmrSTxt
If picAddr.Width > picZoom.Width + 500 Then picZoom.Left = picAddr.Width - picZoom.Width - 120
txtAddr.Width = picAddr.Width - 600 - picZoom.Width - imgWebIco(0).Width
ResizeTabs
End Sub

Private Sub txtAddr_Change()
If txtAddr.Text = "" Then Exit Sub
lstAddr.EnterTxt txtAddr.Text
End Sub

Private Sub txtAddr_Click()
lstAddr.Width = txtAddr.Width
lstAddr.Visible = True
End Sub

Private Sub txtAddr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtAddr_KeyUp(KeyCode As Integer, Shift As Integer)
lstAddr.Visible = txtAddr.Text <> ""
If KeyCode = 13 Then
If lstAddr.SeledSearch = True Then
WebGoTo ReadCon("Search") & lstAddr.Selected
Else
WebGoTo lstAddr.Selected
DoEvents
SaveHistory lstAddr.Selected
End If
Web(NowPage).SetFocus
ElseIf KeyCode = vbKeyUp Then
lstAddr.ChangeSelect 0
txtAddr.SelStart = 0
ElseIf KeyCode = vbKeyDown Then
lstAddr.ChangeSelect 1
txtAddr.SelStart = 0
End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtPass.Text = ReadCon("Password") Then
KeyAscii = 0
tmrSPass.Enabled = True
Web(NowPage).SetFocus
txtPass.Text = ""
End If
End Sub

Private Sub Web_BfrNav2(Index As Integer, ByVal pDisp As Object, Url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If Index = NowPage Then SetOrder imgBtn(2)
End Sub

Private Sub Web_DwnComplete(Index As Integer)
On Error Resume Next
If Index = NowPage Then
txtAddr.Text = Web(Index).LocationURL
Set Url = Web(Index).Docoment
SetOrder imgBtn(1)
End If
End Sub

Private Sub Web_FilDwn(Index As Integer, ByVal Url As String, FileName As String)
If ReadCon("AutoDownload") = 1 Then
Dim cdlg As New clsCdlg, Ex() As String, sFile As String
Ex = Split(FileName, ".")
cdlg.FileName = FileName
cdlg.ShowSave Me.hWnd, Ex(1) & " 文件" & Chr(0) & "*." & Ex(1), "新建下载"
If cdlg.FileName <> "" Then
sFile = Left(cdlg.FileName, Len(cdlg.FileName) - 1)
If right(sFile, Len(Ex(1)) + 1) <> "." & Ex(1) Then sFile = sFile & "." & Ex(1)
If Dir(MyPath & "MagDown.exe") <> "" Then
Shell MyPath & "MagDown.exe " & Url & "@@" & sFile, vbNormalFocus
End If
End If
End If
End Sub

Private Sub Web_GotFocus(Index As Integer)
NowPage = Index
picImg.Visible = False
lstAddr.Visible = False
End Sub

Private Sub Web_NavComplete2(Index As Integer, ByVal pDisp As Object, Url As Variant)
ResizeTabs
Web(NowPage).DoCommand 5, Val(Replace(lblZmVl.Caption, "%", ""))
If Index = NowPage Then RunOnLoadApp
End Sub

Private Sub Web_NewWndw2(Index As Integer, ppDisp As Object, Cancel As Boolean)
AddPage
Set ppDisp = Web(Web.UBound).wObject   '2.新窗口
End Sub

Private Sub Web_ProChange(Index As Integer, ByVal Percent As Long)
On Error Resume Next
Tabs(Index).Progress = Percent * 100
End Sub

Private Sub Web_SetScrLckIcon(Index As Integer, ByVal SecureVal As Long)
If Index = NowPage Then
If Web(Index).SecureVal <> 0 Then SetOrder imgWebIco(1) Else SetOrder imgWebIco(0)
End If
End Sub

Private Sub Web_StaTxtChange(Index As Integer, ByVal Text As String)
If Index = NowPage Then
If Text = "" Or Text = "完成" Then
picSta.Visible = False
Else
lblSta.Caption = Text
picSta.Width = lblSta.Width + 120
picSta.Visible = True
SetOrder picSta
End If
End If
End Sub

Private Sub Web_TtlChange(Index As Integer, ByVal Text As String)
'On Error GoTo TtlErr
If ReadCon("UseFilter") = 1 Then FilterWeb Text
If Text = "广告" Then Tabs_TabClose (Index): Me.ZOrder 0: Exit Sub
Tabs(Index).Caption = Text
If Index = NowPage Then SetCaption Text & " - Magnifier"
Exit Sub
'TtlErr:
'If Err.Number = 340 Then Tabs(GetTabIndex(Index)).Caption = Text: Resume Next
End Sub

Sub DoMouse(SS As String)
Select Case SS
Case "左"
Web(NowPage).DoCommand 0
Case "右"
Web(NowPage).DoCommand 1
Case "下左"
Tabs_TabClose CInt(NowPage)
Case "下右"
If MouseOpen Then
imgAddWeb_MouseUp 1, 0, 1, 1
MouseOpen = False
End If
Case "下上", "上下"
Web(NowPage).DoCommand 3
End Select
End Sub
