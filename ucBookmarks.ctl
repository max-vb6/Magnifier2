VERSION 5.00
Begin VB.UserControl ucBookmarks 
   BackColor       =   &H00EBEBEB&
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   5670
   ScaleWidth      =   4830
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4815
      TabIndex        =   7
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox txtCap 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   4695
      End
      Begin Magnifier.ucCmdBtn btnAdd 
         Height          =   615
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         Caption         =   "Ìí¼ÓÊéÇ©"
      End
      Begin VB.Label lblWhite 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picBM 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÊéÇ©"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1125
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   0
         Width           =   1680
      End
      Begin VB.Image imgBMIco 
         Height          =   1065
         Left            =   0
         Picture         =   "ucBookmarks.ctx":0000
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.FileListBox filMmb 
      Height          =   270
      Left            =   0
      Pattern         =   "*.mmb"
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar sroBM 
      Height          =   4455
      LargeChange     =   100
      Left            =   4560
      Max             =   100
      SmallChange     =   10
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBMLst 
      Appearance      =   0  'Flat
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   1200
      Width           =   4815
      Begin VB.Image imgDel 
         Height          =   375
         Left            =   3960
         Picture         =   "ucBookmarks.ctx":010C
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BookMark 1"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSel 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Tag             =   "1"
         Top             =   0
         Width           =   4815
      End
   End
End
Attribute VB_Name = "ucBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event BookmarkOpened(bmPath As String)
Event BookmarkAdded(bmCaption As String)

Private Sub btnAdd_Click()
If txtCap.Text = "" Then Beep: Exit Sub
RaiseEvent BookmarkAdded(txtCap.Text)
End Sub

Private Sub imgDel_Click()
On Error Resume Next
If MsgBox("È·ÈÏÉ¾³ýÊéÇ© " & lblItems(Int(lblSel.Tag)).Caption & " Âð£¿", 64 + vbYesNo, "É¾³ýÊéÇ©") = vbYes Then
Kill lblItems(Int(lblSel.Tag)).Tag
GetBookmarks
End If
End Sub

Private Sub lblItems_Click(Index As Integer)
lblSel.Top = lblItems(Index).Top - 120
lblSel.Tag = Index
imgDel.Top = lblSel.Top + (lblSel.Height - imgDel.Height) / 2
End Sub

Private Sub lblItems_DblClick(Index As Integer)
RaiseEvent BookmarkOpened(lblItems(Index).Tag)
End Sub

Private Sub lblSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
With lblSel
Static oy!
If Button = 1 Then
.Top = .Top - oy + Y
imgDel.Top = lblSel.Top + (lblSel.Height - imgDel.Height) / 2
Else
oy = Y
End If
End With
End Sub

Private Sub lblSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblSel.Top > picBMLst.Height Then lblSel.Top = picBMLst.Height - lblSel.Height
lblSel.Top = lblItems((lblSel.Top + lblSel.Height) / lblSel.Height).Top - 120
imgDel.Top = lblSel.Top + (lblSel.Height - imgDel.Height) / 2
RaiseEvent BookmarkOpened(lblItems((lblSel.Top + lblSel.Height) / lblSel.Height).Tag)
End Sub

Private Sub sroBM_Change()
picBMLst.Top = picBM.Height - sroBM.Value
End Sub

Private Sub sroBM_Scroll()
sroBM_Change
End Sub

Private Sub UserControl_Initialize()
imgDel.Top = (lblSel.Height - imgDel.Height) / 2
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 4830
If picSave.Visible = False Then
UserControl.Height = 5670
Else
UserControl.Height = picBM.Height + picSave.Height
End If
End Sub

Sub GetBookmarks()
Dim i As Long
picSave.Visible = False
UserControl_Resize
filMmb.Path = MyPath & "Data\Bookmarks\"
filMmb.Refresh
For i = 0 To lblItems.UBound
If i > 0 Then Unload lblItems(i)
Next i
For i = 1 To filMmb.ListCount
Load lblItems(i)
With lblItems(i)
.Move lblItems(0).Left, lblSel.Height * (i - 1) + 120
.Caption = GetFileName(filMmb.List(i - 1))
.Tag = MyPath & "Data\Bookmarks\" & filMmb.List(i - 1)
.Visible = True
.ZOrder 0
End With
Next i
picBMLst.Height = lblSel.Height * (lblItems.Count - 1)
If picBMLst.Height > UserControl.Height - picBM.Height Then
sroBM.Max = picBMLst.Height - (UserControl.Height - picBM.Height)
sroBM.SmallChange = lblSel.Height
sroBM.LargeChange = lblSel.Height * 2
sroBM.Visible = True
End If
End Sub

Sub ShowSave(sCap As String)
txtCap.Text = sCap
picSave.Visible = True
UserControl_Resize
End Sub
