VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmApp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5460
   Icon            =   "frmApp.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5460
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin SHDocVwCtl.WebBrowser WbHtml 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      ExtentX         =   2355
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim App_hMod As Long, App_hMod2 As Long  'WebBrowser Style Setting

Private Sub Form_Initialize()
Dim iccex As tagInitCommonControlsEx
iccex.lngSize = LenB(iccex)
iccex.lngICC = ICC_USEREX_CLASSES
InitCommonControlsEx iccex
App_hMod = LoadLibrary("shell32.dll")
App_hMod2 = LoadLibrary("explorer.exe")
End Sub

Private Sub Form_Load()
WbHtml.Silent = True
WbHtml.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
WbHtml.Move -23, -23, Me.ScaleWidth + 46, Me.ScaleHeight + 46
End Sub

Private Sub Form_Terminate()
If App_hMod Then FreeLibrary App_hMod
If App_hMod2 Then FreeLibrary App_hMod2
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmSHide.SC(CInt(Me.Tag)).AddCode "Sub TheMagLetUEnd" & vbCrLf & "AppControl.EndApp" & vbCrLf & "End Sub"
frmSHide.SC(CInt(Me.Tag)).Run "TheMagLetUEnd"
End Sub

Private Sub WbHtml_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
End Sub

Private Sub WbHtml_TitleChange(ByVal Text As String)
Me.Caption = Text
End Sub

Sub UnloadForm()
Unload Me
End Sub

Sub SetHtml(sHtmPath As String)
On Error Resume Next
WbHtml.Navigate sHtmPath
End Sub
