VERSION 5.00
Begin VB.UserControl ucDownload 
   BackColor       =   &H00EBEBEB&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ScaleHeight     =   975
   ScaleWidth      =   4455
   Begin VB.PictureBox picPro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   960
      ScaleHeight     =   135
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      Begin VB.Label lblPro 
         BackColor       =   &H00008000&
         Height          =   255
         Left            =   -10
         TabIndex        =   5
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3360
      ScaleHeight     =   615
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      Begin VB.Image imgCP 
         Height          =   375
         Index           =   1
         Left            =   0
         Picture         =   "ucDownload.ctx":0000
         ToolTipText     =   "暂停下载"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgCP 
         Height          =   375
         Index           =   0
         Left            =   0
         Picture         =   "ucDownload.ctx":0098
         ToolTipText     =   "继续下载"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgDel 
         Height          =   375
         Left            =   480
         Picture         =   "ucDownload.ctx":011C
         ToolTipText     =   "从列表中删除这个下载"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Label lblProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已下载 *%"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   840
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FileName"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "ucDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event DwnRemoved()
Dim WithEvents Downloader As CFileDownload
Attribute Downloader.VB_VarHelpID = -1
Dim dUrl As String, dPath As String

Sub NewDownload(sUrl As String, sPath As String)
If LCase(Left(sUrl, 6)) = "ftp://" Then Exit Sub: imgDel_Click
dUrl = sUrl
dPath = sPath
imgCP_Click 0
End Sub

Sub StopDownload()
Downloader.AbortDownloading
End Sub

Private Sub Downloader_OnProgress(ByVal lProgress As Long, ByVal lMaxProgress As Long, ByVal lStatusCode As Long, ByVal sStatusText As String)
If lMaxProgress = 0 Then
lblProg.Caption = "已下载 0%"
lblPro.Width = 0
Else
lblProg.Caption = "已下载 " & Int(lProgress / lMaxProgress * 100) & "%"
lblPro.Width = Int(lProgress / lMaxProgress * picPro.ScaleWidth) + 10
If Int(lProgress / lMaxProgress * 100) = 100 Then
imgCP(0).Visible = False
imgCP(1).Visible = False
imgDel.Left = 0
picOpt.Width = imgDel.Width + 240
UserControl_Resize
lblProg.Caption = "下载完成"
picPro.Visible = False
GetFileIcon dPath, picIcon
If Right(lblCap.Caption, 5) = ".mapx" Then Shell App.Path & "\Magnifier.exe " & dPath, vbNormalFocus
End If
End If
UserControl_Resize
End Sub

Private Sub imgCP_Click(Index As Integer)
imgCP(Index).Visible = False
imgCP(Abs(Index - 1)).Visible = True
Select Case Index
Case 0
lblCap.Caption = GetFileName(dPath, True)
GetFileIcon dPath, picIcon
DoEvents
Downloader.StartDownloading dUrl, dPath
Case 1
Downloader.AbortDownloading
End Select
End Sub

Private Sub imgDel_Click()
Downloader.AbortDownloading
RaiseEvent DwnRemoved
End Sub

Private Sub lblCap_DblClick()
If picPro.Visible = False Then ShellExecute 0, "open", dPath, 0, 0, 1
End Sub

Private Sub picIcon_DblClick()
lblCap_DblClick
End Sub

Private Sub UserControl_Initialize()
Set Downloader = New CFileDownload
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 975
picOpt.Left = UserControl.Width - picOpt.Width
picOpt.Top = (UserControl.Height - picOpt.Height) / 2
lblProg.Left = picOpt.Left - lblProg.Width - 120
picPro.Top = lblProg.Top + (lblProg.Height - picPro.Height) / 2
picPro.Width = UserControl.Width - picPro.Left - picOpt.Width - lblProg.Width - 240
End Sub

