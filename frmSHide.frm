VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmSHide 
   BorderStyle     =   0  'None
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSScriptControlCtl.ScriptControl SC 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Image imgSH 
      Height          =   615
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgSH 
      Height          =   330
      Index           =   0
      Left            =   0
      Picture         =   "frmSHide.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgSH 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "frmSHide.frx":2BEA
      Top             =   0
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Image imgSH 
      Height          =   390
      Index           =   2
      Left            =   0
      Picture         =   "frmSHide.frx":5D00
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "frmSHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = ReadCon("SHKey") Then
frmMain.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
imgSH(3).Picture = LoadPicture(ReadCon("SHCustom"))
With imgSH(Val(ReadCon("SuperHide")))
Me.Picture = .Picture
Me.Move Screen.Width - .Width, Screen.Height - GetTaskbarHeight * Screen.TwipsPerPixelY - .Height, .Width, .Height
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 5
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
End Sub
