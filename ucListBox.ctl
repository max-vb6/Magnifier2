VERSION 5.00
Begin VB.UserControl ucListBox 
   BackColor       =   &H00FEDCC0&
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   1935
   ScaleWidth      =   4815
   Begin VB.Label lblHis 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
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
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image imgHis 
      Height          =   285
      Left            =   120
      Picture         =   "ucListBox.ctx":0000
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "找到 * 个关于 * 的应用"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1830
   End
   Begin VB.Label lblGoto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转到 *"
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
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgGoto 
      Height          =   285
      Left            =   120
      Picture         =   "ucListBox.ctx":00A4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lblSrch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "搜索 *"
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgSrch 
      Height          =   285
      Left            =   120
      Picture         =   "ucListBox.ctx":017D
      Top             =   600
      Width           =   285
   End
   Begin VB.Label lblSel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFCB9B&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "ucListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event MouseClicked()
Dim sHisLst() As String, sAppLst() As String, strSel As String, SelSrch As Boolean

Public Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get SelColor() As Long
    SelColor = lblSel.BackColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Long)
    lblSel.BackColor() = New_SelColor
    PropertyChanged "SelColor"
End Property

Public Property Get ForeColor() As Long
    ForeColor = lblGoto.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    lblGoto.ForeColor = New_ForeColor
    lblSrch.ForeColor = New_ForeColor
    lblApp.ForeColor = New_ForeColor
    lblHis(0).ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Selected() As String
    Selected = strSel
End Property

Public Property Get SeledSearch() As Boolean
    SeledSearch = SelSrch
End Property

Private Sub EmptyArrays()
Dim i As Long
For i = 0 To UBound(sAppLst)
sAppLst(i) = ""
Next i
For i = 0 To UBound(sHisLst)
sHisLst(i) = ""
If i + 1 <= 10 Then Unload lblHis(i + 1)
Next i
End Sub

Private Sub lblSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
With lblSel
Static oy!
If Button = 1 Then
.Top = .Top - oy + Y
Else
oy = Y
End If
End With
End Sub

Private Sub lblSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblSel.Top < 0 Then lblSel.Top = 0
If lblSel.Top + lblSel.Height > UserControl.Height Then lblSel.Top = UserControl.Height - lblSel.Height
If CtrlInCtrl(lblGoto, lblSel) Or CtrlInCtrl(lblApp, lblSel) Then
strSel = lblGoto.Tag: SelSrch = False
ElseIf CtrlInCtrl(lblSrch, lblSel) Then
strSel = lblGoto.Tag: SelSrch = True
Else
strSel = lblHis((lblSel.Top - lblSel.Height * 2) / lblSel.Height).Caption
End If
lblSel.Top = 0
RaiseEvent MouseClicked
End Sub

Private Sub UserControl_Resize()
lblSel.Width = UserControl.ScaleWidth
End Sub

Sub EnterTxt(Txt As String)
On Error Resume Next
EmptyArrays
UserControl.Height = lblSel.Height * 3
lblGoto.Caption = "转到 " & Txt
lblGoto.ToolTipText = lblGoto.Caption
lblGoto.Tag = Txt
lblSrch.Caption = "搜索 " & Txt
lblSrch.ToolTipText = lblSrch.Caption
If InStr(Txt, ".") = 0 And InStr(Txt, "/") = 0 And InStr(Txt, ":") = 0 Then
lblSel.Top = lblSel.Height
SelSrch = True
Else
lblSel.Top = 0
SelSrch = False
End If
strSel = Txt

If SearchApp(Txt) = "" Then
lblApp.Caption = "未找到与 " & Txt & " 相关的应用"
Else
sAppLst = Split(SearchApp(Txt), "@@")
lblApp.Caption = "找到 " & UBound(sAppLst) & " 个关于 " & Txt & " 的应用"
lblApp.ToolTipText = lblApp.Caption
End If
If SearchHistory(Txt) <> "" Then
sHisLst = Split(SearchHistory(Txt), "@@")
Dim i As Long, lNum As Long
lNum = UBound(sHisLst) + 1
If lNum > 10 Then lNum = 10
For i = 1 To lNum
Load lblHis(i)
lblHis(i).Visible = True
lblHis(i).ZOrder 0
lblHis(i).Move lblHis(0).Left, lblSel.Height * 3 + 120 + (240 + lblHis(0).Height) * (i - 1)
lblHis(i).Caption = sHisLst(i - 1)
lblHis(i).ToolTipText = lblHis(i).Caption
Next i
UserControl.Height = lblSel.Height * (2 + lNum)
lblSel.Top = lblSel.Height * 3
strSel = lblHis(1).Caption
SelSrch = False
End If
End Sub

Sub ChangeSelect(UpOrDown As Long)         'Up = 0 ; Down = Else
If UpOrDown = 0 Then
If lblSel.Top > 0 Then lblSel.Top = lblSel.Top - lblSel.Height
Else
If lblSel.Top + lblSel.Height < UserControl.ScaleHeight Then lblSel.Top = lblSel.Top + lblSel.Height
End If
If lblSel.Top = 0 Or lblSel.Top = lblSel.Height Or lblSel.Top = lblSel.Height * 2 Then
strSel = lblGoto.Tag
Else
strSel = lblHis((lblSel.Top - lblSel.Height * 2) / lblSel.Height).Caption
End If
If lblSel.Top = lblSel.Height Then
SelSrch = True
Else
SelSrch = False
End If
End Sub
