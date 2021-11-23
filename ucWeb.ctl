VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.UserControl ucWeb 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ScaleHeight     =   3615
   ScaleWidth      =   4470
   Begin VB.PictureBox picNav 
      BackColor       =   &H00FEDCC0&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   3315
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3315
      Begin VB.PictureBox picAddWeb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   2160
         ScaleHeight     =   2865
         ScaleWidth      =   5025
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtAddUrl 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Î¢ÈíÑÅºÚ"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   4575
         End
         Begin Magnifier.ucCmdBtn cmdCnl 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2040
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   873
            Caption         =   "È¡Ïû"
         End
         Begin Magnifier.ucCmdBtn cmdAddTxt 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1560
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   873
            Caption         =   "Ìí¼Ó"
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ìí¼ÓÒ³ÃæÖÁµ¼º½"
            BeginProperty Font 
               Name            =   "Î¢ÈíÑÅºÚ"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Width           =   2100
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   0
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   0
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Left            =   0
            TabIndex        =   28
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   1
         Left            =   3240
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   0
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   2
         Left            =   6480
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   0
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   3
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   2280
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   3
            Left            =   0
            TabIndex        =   22
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   4
         Left            =   3240
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   2280
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   4
            Left            =   0
            TabIndex        =   20
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   5
         Left            =   6480
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   2280
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   5
            Left            =   0
            TabIndex        =   18
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   6
         Left            =   0
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   6
            Left            =   0
            TabIndex        =   16
            ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   7
         Left            =   3240
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   4560
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   7
            Left            =   0
            TabIndex        =   14
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.PictureBox picWeb 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFCB9B&
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   8
         Left            =   6480
         ScaleHeight     =   1905
         ScaleWidth      =   2865
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "µã»÷ÓÒ¼ü±à¼­µ¼º½"
         Top             =   4560
         Width           =   2895
         Begin VB.Label lblWeb 
            Alignment       =   2  'Center
            BackColor       =   &H00FFCB9B&
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
            Index           =   8
            Left            =   0
            TabIndex        =   12
            Top             =   1440
            Width           =   2895
         End
      End
      Begin VB.Image imgBG 
         Height          =   2055
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox picError 
      BackColor       =   &H00FEDCC0&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   3135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label lblShow 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   2160
         Width           =   150
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ôã¸â£¬³öÁËµãÐ¡ÎÊÌâ..."
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   1080
         Width           =   4230
      End
      Begin VB.Image imgWhale 
         Height          =   2955
         Left            =   0
         Picture         =   "ucWeb.ctx":0000
         Top             =   0
         Width           =   3510
      End
   End
   Begin SHDocVwCtl.WebBrowser Wb 
      Height          =   2655
      Left            =   -23
      TabIndex        =   0
      Top             =   -23
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
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
      Location        =   "http:///"
   End
   Begin VB.PictureBox picWebImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   94
      TabIndex        =   4
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "ucWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_hMod As Long, m_hMod2 As Long  'WebBrowser Style Setting
Dim WebError As Boolean, MobSwi As Boolean, IsNav As Boolean, NowTitle As String, PageSecure As Long

Event StaTxtChange(ByVal Text As String)
Event BfrNav2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Event DocComplete(ByVal pDisp As Object, URL As Variant)
Event DwnComplete()
Event FilDwn(ByVal URL As String, FileName As String)
Event NavComplete2(ByVal pDisp As Object, URL As Variant)
Event ProChange(ByVal Percent As Long)
Event NewWndw2(ppDisp As Object, Cancel As Boolean)
Event TtlChange(ByVal Text As String)
Event SetScrLckIcon(ByVal SecureVal As Long)

Private Sub cmdAddTxt_Click()
If txtAddUrl.Text <> "" Then
lblWeb(picAddWeb.Tag).Tag = txtAddUrl.Text
lblWeb(picAddWeb.Tag).Caption = txtAddUrl.Text
picWeb(picAddWeb.Tag).Tag = ""
picWeb(picAddWeb.Tag).Picture = LoadPicture()
WriteString "Nav", "n" & picAddWeb.Tag + 1, txtAddUrl.Text & "@@" & txtAddUrl.Text & "@@", MyPath & "Data\MagNav.ini"
Else
lblWeb(picAddWeb.Tag).Tag = ""
lblWeb(picAddWeb.Tag).Caption = ""
picWeb(picAddWeb.Tag).Tag = ""
picWeb(picAddWeb.Tag).Picture = LoadPicture()
WriteString "Nav", "n" & picAddWeb.Tag + 1, "", MyPath & "Data\MagNav.ini"
End If
picAddWeb.Visible = False
End Sub

Private Sub cmdCnl_Click()
picAddWeb.Visible = False
End Sub

Private Sub lblWeb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picWeb_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub picWeb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If lblWeb(Index).Tag <> "" Then
Wb.Navigate lblWeb(Index).Tag
IsNav = True
Else
GoTo Add
End If
Else
Add:
picAddWeb.Tag = Index
picAddWeb.Visible = True
End If
End Sub

Private Sub UserControl_Initialize()
Wb.Silent = True
Dim iccex As tagInitCommonControlsEx
iccex.lngSize = LenB(iccex)
iccex.lngICC = ICC_USEREX_CLASSES
InitCommonControlsEx iccex
m_hMod = LoadLibrary("shell32.dll")
m_hMod2 = LoadLibrary("explorer.exe")
End Sub

Private Sub UserControl_Resize()
Wb.Move -23, -23, UserControl.ScaleWidth + 46, UserControl.ScaleHeight + 46
picNav.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
imgBG.Move 0, 0, picNav.Width, picNav.Height
picError.Move 0, 0, picNav.Width, picNav.Height
picAddWeb.Move (picNav.Width - picAddWeb.Width) / 2, (picNav.Height - picAddWeb.Height) / 2
'===MoveNavWeb===
Dim I As Long
For I = 0 To 2
picWeb(I).Move (picNav.Width / 3 - picWeb(I).Width) / 2 + picNav.Width / 3 * I, (picNav.Height / 3 - picWeb(I).Height) / 2 + picNav.Height / 3 * 0
Next I
For I = 3 To 5
picWeb(I).Move (picNav.Width / 3 - picWeb(I).Width) / 2 + picNav.Width / 3 * (I - 3), (picNav.Height / 3 - picWeb(I).Height) / 2 + picNav.Height / 3 * 1
Next I
For I = 6 To 8
picWeb(I).Move (picNav.Width / 3 - picWeb(I).Width) / 2 + picNav.Width / 3 * (I - 6), (picNav.Height / 3 - picWeb(I).Height) / 2 + picNav.Height / 3 * 2
Next I
'======End=======
imgWhale.Move picNav.Width - imgWhale.Width, picNav.Height - imgWhale.Height
lblShow(0).Move (picNav.Width - lblShow(0).Width) / 2
lblShow(1).Move (picNav.Width - lblShow(1).Width) / 2
End Sub

Private Sub UserControl_Terminate()
If m_hMod Then FreeLibrary m_hMod
If m_hMod2 Then FreeLibrary m_hMod2
End Sub

Private Sub Wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
RaiseEvent BfrNav2(pDisp, URL, flags, TargetFrameName, PostData, Headers, Cancel) '
If URL = "about:nav" Then picNav.Visible = True: LoadNav
If picError.Visible = True Then picError.Visible = False: WebError = False
If ReadCon("MobileMode") <> 0 Then
If MobSwi = True Then Exit Sub
MobSwi = True
Wb.Navigate URL, , , , "User-Agent: " & ReadCon("UA")
Cancel = True
End If
PageSecure = 0
End Sub

Private Sub Wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
RaiseEvent DocComplete(pDisp, URL)
If URL <> "about:nav" Then picNav.Visible = False: picAddWeb.Visible = False
If IsNav And Not (WebError) Then IsNav = False: GetNavPic
MobSwi = False
End Sub

Private Sub Wb_DownloadComplete()
RaiseEvent DwnComplete
End Sub

Private Sub Wb_FileDownload(ByVal ActiveDocument As Boolean, Cancel As Boolean)
Dim sFile As String, URL As String
On Error GoTo errH:
URL = Wb.LocationURL
If Wb.Document.activeElement.href <> "" Then URL = Wb.Document.activeElement.href
sFile = DownloadUrlToName(URL)
RaiseEvent FilDwn(URL, sFile)
errH:
End Sub

Private Sub Wb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
RaiseEvent NavComplete2(pDisp, URL)
End Sub

Private Sub Wb_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
If InStr(LCase(URL), "java") = 0 And InStr(LCase(URL), "eiv.baidu.com") = 0 Then
lblShow(1).Caption = "Õâ¸öÒ³ÃæÔÝÊ±ÎÞ·¨ÏÔÊ¾£¬Äú¿ÉÒÔ³¢ÊÔË¢ÐÂÒ³Ãæ" & vbCrLf & "»òÕß´ÓËÑË÷ÒýÇæ¼ìË÷¸ÃÍøÕ¾µÄÏà¹ØÐÅÏ¢" & vbCrLf & vbCrLf & "´íÎó´úÂë " & StatusCode
picError.BackColor = ThmClr.MainColor
picError.Visible = True
picError.ZOrder 0
WebError = True
End If
End Sub

Private Sub Wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
RaiseEvent NewWndw2(ppDisp, Cancel)
End Sub

Private Sub Wb_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
RaiseEvent ProChange(Progress / ProgressMax)
End Sub

Private Sub Wb_SetSecureLockIcon(ByVal SecureLockIcon As Long)
If PageSecure = 0 And SecureLockIcon <> 0 Then PageSecure = SecureLockIcon
RaiseEvent SetScrLckIcon(SecureLockIcon)
End Sub

Private Sub Wb_StatusTextChange(ByVal Text As String)
On Error Resume Next
RaiseEvent StaTxtChange(Text)
End Sub

Private Sub Wb_TitleChange(ByVal Text As String)
If WebError Then
NowTitle = "Magnifier ÎÞ·¨ÔØÈëÒ³Ãæ"
Else
NowTitle = Text
End If
If Text = "about:nav" Then NowTitle = "Magnifier µ¼º½"
RaiseEvent TtlChange(NowTitle)
End Sub

Public Property Get LocationName() As String
LocationName = NowTitle
End Property

Public Property Get LocationURL() As String
LocationURL = Wb.LocationURL
End Property

Public Property Get wObject() As Object
Set wObject = Wb.Object
End Property

Public Property Let wObject(ByVal New_wObject As Object)
Set Wb.Object = New_wObject
PropertyChanged "wObject"
End Property

Public Property Get Docoment() As HTMLDocument
On Error Resume Next
Set Docoment = Wb.Document
End Property

Public Property Let Document(ByVal New_Document As HTMLDocument)
Set Wb.Document = New_Document
PropertyChanged "Document"
End Property

Public Property Get SecureVal() As Long
SecureVal = PageSecure
End Property

Sub GoURL(ByVal URL As String, Optional flags As Variant = "", Optional TargetFrameName As Variant = "", Optional PostData As Variant = "", Optional Headers As Variant = "")
On Error Resume Next
Wb.Navigate URL, flags, TargetFrameName, PostData, Headers
End Sub

Sub DoCommand(ByVal lCommand As Long, Optional flags As Long = 100)
On Error Resume Next
With Wb
Select Case lCommand
Case 0
.GoBack
Case 1
.GoForward
Case 2
.Stop
Case 3
.Refresh
Case 4
.GoHome
Case 5
.Document.body.Style.Zoom = CStr(flags) & "%"
Case 6
.ExecWB 4, 1    'Save Page
Case 7
.ExecWB 6, 1    'Print Page
End Select
End With
End Sub

Private Sub LoadNav()
On Error Resume Next
Dim navstr As String, strs() As String, I As Long
For I = 0 To 8
picWeb(I).BackColor = ThmClr.ToolBar
lblWeb(I).BackColor = picWeb(I).BackColor
navstr = ReadString("Nav", "n" & I + 1, MyPath & "Data\MagNav.ini")
If navstr <> "" Then
strs = Split(navstr, "@@")
lblWeb(I).Tag = strs(0)
lblWeb(I).Caption = strs(1)
lblWeb(I).ToolTipText = strs(1)
picWeb(I).Tag = strs(2)
picWeb(I).PaintPicture LoadPicture(MyPath & picWeb(I).Tag), 0, 0, picWeb(I).Width, picWeb(I).Height
Else
picWeb(I).Tag = ""
picWeb(I).Picture = LoadPicture()
lblWeb(I).Tag = ""
lblWeb(I).Caption = ""
lblWeb(I).ToolTipText = ""
End If
Next I
picNav.BackColor = ThmClr.MainColor
If Dir(GetThmFolder & "NavBG.jpg") <> "" Then imgBG.Picture = LoadPicture(GetThmFolder & "NavBG.jpg")
End Sub

Private Sub GetNavPic()
On Error Resume Next
Dim I As Long
For I = 0 To 8
If Replace(Replace(lblWeb(I).Tag, "http:", ""), "/", "") = Replace(Replace(Wb.LocationURL, "http:", ""), "/", "") Then
picWeb(I).Tag = "Data\NavPic\n" & I + 1 & ".bmp"
SavePicture GetWebImg, MyPath & picWeb(I).Tag
WriteString "Nav", "n" & I + 1, Wb.LocationURL & "@@" & NowTitle & "@@" & picWeb(I).Tag, MyPath & "Data\MagNav.ini"
Exit For
End If
Next I
End Sub

Function GetWebImg() As stdole.StdPicture
On Error Resume Next
With picWebImg
.Width = Wb.Width - 20 * Screen.TwipsPerPixelX          '¼õ20ÏñËØ£¬±ÜÃâ¹ö¶¯Ìõ
.Height = Wb.Height - 20 * Screen.TwipsPerPixelX
Set .Picture = LoadPicture()        'Çå³ý¾ÉµÄÍ¼Ïñ
.AutoRedraw = True
PrintWindow UserControl.hWnd, .hDC, 0
DoEvents
Set GetWebImg = .Image
End With
End Function
