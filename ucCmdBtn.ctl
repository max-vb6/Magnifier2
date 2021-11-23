VERSION 5.00
Begin VB.UserControl ucCmdBtn 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   ScaleHeight     =   1005
   ScaleWidth      =   3645
   Begin VB.Timer tmrClr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblCaption 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   60
   End
End
Attribute VB_Name = "ucCmdBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ColorTo As Long
Event Click()
Event DblClick()

Private Sub lblCaption_Click()
UserControl_Click
End Sub

Private Sub lblCaption_DblClick()
UserControl_DblClick
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub tmrClr_Timer()
With UserControl
Dim nc&
nc = (.BackColor And &HFF) Mod 256
If ColorTo > nc Then
.BackColor = RGB(nc + 5, nc + 5, nc + 5)
If nc >= ColorTo Then .BackColor = RGB(ColorTo, ColorTo, ColorTo): tmrClr.Enabled = False
Else
.BackColor = RGB(nc - 5, nc - 5, nc - 5)
If nc <= ColorTo Then .BackColor = RGB(ColorTo, ColorTo, ColorTo): tmrClr.Enabled = False
End If
End With
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
UserControl.BackColor = RGB(235, 235, 235)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ColorTo = 224
tmrClr.Enabled = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (0 <= X) And (X <= UserControl.Width) And (0 <= Y) And (Y <= UserControl.Height) Then
ColorTo = 255
SetCapture UserControl.hwnd
Else
ColorTo = 235
ReleaseCapture
End If
tmrClr.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ColorTo = 255
tmrClr.Enabled = True
End Sub

Private Sub UserControl_Paint()
UserControl_Initialize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
lblCaption.Caption = PropBag.ReadProperty("Caption", "")
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
lblCaption.Move (UserControl.ScaleWidth - lblCaption.Width) / 2, (UserControl.ScaleHeight - lblCaption.Height) / 2
End Sub

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal nCap As String)
    PropertyChanged "Caption"
    lblCaption.Caption = nCap
    UserControl_Resize
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", lblCaption.Caption
End Sub
