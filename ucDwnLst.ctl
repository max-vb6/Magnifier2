VERSION 5.00
Begin VB.UserControl ucDwnLst 
   BackColor       =   &H00EBEBEB&
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   5670
   ScaleWidth      =   4830
   Begin VB.VScrollBar sroDwn 
      Height          =   4455
      LargeChange     =   100
      Left            =   4560
      Max             =   100
      SmallChange     =   10
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picDLst 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
      Begin MagDown.ucDownload Dwner 
         Height          =   975
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1720
      End
   End
   Begin VB.PictureBox picDwn 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Label lblClear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "清空列表"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   405
         Left            =   3480
         TabIndex        =   4
         Top             =   720
         Width           =   1200
      End
      Begin VB.Image imgDwIco 
         Height          =   1065
         Left            =   120
         Picture         =   "ucDwnLst.ctx":0000
         Top             =   120
         Width           =   1200
      End
   End
End
Attribute VB_Name = "ucDwnLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Sub Dwner_DwnRemoved(Index As Integer)
Unload Dwner(Index)
ReAvg
End Sub

Private Sub lblClear_Click()
If Dwner.Count = 1 Then Exit Sub
If MsgBox("确定要清空下载列表吗？", 32 + vbYesNo, "清空") = vbNo Then Exit Sub
On Error GoTo errRe
Dim i As Long
For i = 1 To Dwner.UBound
Dwner(i).StopDownload
Unload Dwner(i)
Next i
Exit Sub
errRe:
i = i + 1
Resume
End Sub

Private Sub sroDwn_Change()
picDLst.Top = picDwn.Height - sroDwn.Value
End Sub

Private Sub sroDwn_Scroll()
sroDwn_Change
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 4830
UserControl.Height = 5670
End Sub

Sub AddNewDwn(sUrl As String, sPath As String)
Load Dwner(Dwner.UBound + 1)
With Dwner(Dwner.UBound)
.Move 0, Dwner(0).Height * (Dwner.Count - 2)
.Visible = True
picDLst.Height = (Dwner.Count - 1) * Dwner(0).Height
If picDLst.Height > UserControl.Height - picDwn.Height Then
sroDwn.LargeChange = Dwner(0).Height * 2
sroDwn.SmallChange = Dwner(0).Height
sroDwn.Visible = True
Else
sroDwn.Visible = False
End If
.NewDownload sUrl, sPath
End With
End Sub

Private Sub ReAvg()
On Error GoTo errRe
Dim i As Long, lTop As Long
lTop = 0
For i = 1 To Dwner.UBound
Dwner(i).Top = lTop
lTop = lTop + Dwner(0).Height
Next i
picDLst.Height = (Dwner.Count - 1) * Dwner(0).Height
If picDLst.Height > UserControl.Height - picDwn.Height Then
sroDwn.LargeChange = Dwner(0).Height * 2
sroDwn.SmallChange = Dwner(0).Height
sroDwn.Visible = True
Else
sroDwn.Visible = False
End If
Exit Sub
errRe:
i = i + 1
Resume
End Sub
