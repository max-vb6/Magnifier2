VERSION 5.00
Begin VB.UserControl ucTabs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   ScaleHeight     =   1665
   ScaleWidth      =   3225
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   0
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H00FDBD93&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   2520
      Begin VB.PictureBox picCls 
         BackColor       =   &H00FDBD93&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         Picture         =   "ucTabs.ctx":0000
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   60
      End
      Begin VB.Label lblPro 
         BackColor       =   &H00FEDCC0&
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Tag             =   "0"
         Top             =   380
         Visible         =   0   'False
         Width           =   15
      End
   End
End
Attribute VB_Name = "ucTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim tWidth As Single
Event TabClose()
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub lblCap_Click()
picTab_Click
End Sub

Private Sub lblCap_DblClick()
picTab_DblClick
End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblPro_Click()
picTab_Click
End Sub

Private Sub lblPro_DblClick()
picTab_DblClick
End Sub

Private Sub lblPro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblPro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblPro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picTab_MouseUp Button, Shift, X, Y
End Sub

Private Sub picCls_Click()
RaiseEvent TabClose
End Sub

Private Sub picTab_Click()
RaiseEvent Click
End Sub

Private Sub picTab_DblClick()
RaiseEvent DblClick
End Sub

Private Sub picTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub tmrSpeed_Timer()
SpeedLessMove lblPro, 0, lblPro.Top, tWidth, lblPro.Height, 6, tmrSpeed
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
If ReadString("Clr", "WhiteFont", MyPath & "Themes\" & ReadThm("Theme")) = 1 Then lblCap.ForeColor = vbWhite
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Height = picTab.Height
picTab.Width = UserControl.Width
picCls.Move picTab.ScaleWidth - picCls.Width, (picTab.ScaleHeight - picCls.Height) / 2
lblCap.Move lblCap.Left, (picTab.ScaleHeight - lblCap.Height) / 2
tmrSpeed.Enabled = False
lblPro.Move 0, lblPro.Top, Int(picTab.ScaleWidth / 100 * lblPro.Tag)
End Sub

Public Property Get BackColor() As Long
    BackColor = picTab.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    picTab.BackColor() = New_BackColor
    picCls.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
    Caption = lblCap.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCap.Caption() = New_Caption
    lblCap.ToolTipText = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Progress() As Integer
    Progress = lblPro.Tag
End Property

Public Property Let Progress(ByVal New_Progress As Integer)
    lblPro.Tag() = New_Progress
    tWidth = Int(picTab.ScaleWidth / 100 * New_Progress)
    If New_Progress <> 0 Then
    tmrSpeed.Enabled = True
    Else
    lblPro.Width = 0
    lblPro.Tag = 0
    End If
    lblPro.Visible = New_Progress <> 0
    PropertyChanged "Progress"
End Property

Public Property Get ProColor() As Long
    ProColor = lblPro.BackColor
End Property

Public Property Let ProColor(ByVal New_ProColor As Long)
    lblPro.BackColor() = New_ProColor
    PropertyChanged "ProColor"
End Property
