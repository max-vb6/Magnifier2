VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "���� Magnifier ��ҳ��������ڶ�����"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrSTxts 
      Interval        =   1
      Left            =   4080
      Top             =   0
   End
   Begin Magnifier.ucCmdBtn cmdOK 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "ȷ��"
   End
   Begin VB.PictureBox picFrm 
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   8055
      TabIndex        =   2
      Top             =   480
      Width           =   8055
      Begin VB.Image imgAbout 
         Height          =   2100
         Left            =   720
         Picture         =   "frmAbout.frx":000C
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.PictureBox picTxts 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   8055
      TabIndex        =   7
      Top             =   4080
      Width           =   8055
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���ĵ����벻��������ѵĴ���֧�֣������Ŷ��ڴ˸�л���ǣ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   8040
      End
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2012��6�£�Magnifier ���������ĵڶ�λ��Ա�������������ҡ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   -15
         TabIndex        =   9
         Top             =   0
         Width           =   8070
      End
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Magnifier2 �̳��˵�һ������������ص㣬������Trident�ں��ϵ���ഴ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   8040
      End
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   1320
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   5760
      Width           =   720
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ٷ���վ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   240
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   5760
      Width           =   720
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩ���� (C) 2010-2012 MaxXSoft ��������. ��������Ȩ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1507
      TabIndex        =   6
      Top             =   3480
      Width           =   5040
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�汾: 1.10.34 Beta3 (�������԰�)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Index           =   1
      Left            =   4110
      TabIndex        =   5
      Top             =   3000
      Width           =   3105
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magnifier ��ҳ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Index           =   0
      Left            =   877
      TabIndex        =   4
      Top             =   3000
      Width           =   2100
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���� Magnifier ��ҳ��������ڶ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   75
      Width           =   3630
   End
   Begin VB.Image imgCtrl 
      Height          =   480
      Left            =   7560
      Picture         =   "frmAbout.frx":2D460
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblCtrl 
      BackColor       =   &H000000C0&
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
picTxts.Left = Me.ScaleWidth
picFrm.Left = -picFrm.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub imgCtrl_Click()
Unload Me
End Sub

Private Sub imgCtrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCtrl.Visible = True
End Sub

Private Sub imgCtrl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCtrl.Visible = False
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblWeb_Click(Index As Integer)
frmMain.AddPage
If Index = 0 Then
frmMain.WebGoTo "http://maxxsoft.net"
Else
frmMain.WebGoTo "http://tieba.baidu.com/f?kw=maxxsoft"
End If
Unload Me
End Sub

Private Sub tmrSTxts_Timer()
SpeedLessMove picTxts, 0, 4080, picTxts.Width, picTxts.Height, 10, tmrSTxts
picFrm.Left = -picTxts.Left
End Sub
