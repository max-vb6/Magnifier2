Attribute VB_Name = "Main"
Public Url As HTMLDocument, MouseOpen As Boolean

'==Alpha==
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT   As Long = &H20&
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
'==End==

'==Other==
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Const ICC_USEREX_CLASSES = &H200
Public Type tagInitCommonControlsEx
lngSize As Long
lngICC As Long
End Type

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Const WM_PRINT As Long = &H317
Public Const PRF_NONCLIENT = &H2
Public Const PRF_CLIENT = &H4
Public Const PRF_ERASEBKGND = &H8
Public Const PRF_CHILDREN = &H10
Public Const PRF_OWNED = &H20

Public Declare Function PrintWindow Lib "user32" (ByVal SrcHwnd As Long, ByVal DesHDC As Long, ByVal uFlag As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_GETMINMAXINFO = &H24
Public procOld As Long
Public udtMMI As MINMAXINFO
Public Type MINMAXINFO
ptReserved As POINTAPI
ptMaxSize As POINTAPI
ptMaxPosition As POINTAPI
ptMinTrackSize As POINTAPI
ptMaxTrackSize As POINTAPI
End Type
'==End==

Public Function SpeedLessMove(mObj As Control, mToLeft As Single, mToTop As Single, mToWidth As Single, mToHeight As Single, Speed As Long, cTimer As Timer, Optional SetTag As String) As Long
Dim ml As Single, mt As Single, mw As Single, mh As Single
On Error Resume Next
If mObj.Left < mToLeft Then
    ml = mToLeft - mObj.Left
    ml = ml / Speed
    mObj.Left = mObj.Left + ml
ElseIf mObj.Left > mToLeft Then
    ml = mObj.Left - mToLeft
    ml = ml / Speed
    mObj.Left = mObj.Left - ml
End If

If mObj.Top < mToTop Then
    mt = mToTop - mObj.Top
    mt = mt / Speed
    mObj.Top = mObj.Top + mt
ElseIf mObj.Top > mToTop Then
    mt = mObj.Top - mToTop
    mt = mt / Speed
    mObj.Top = mObj.Top - mt
End If

If mObj.Width < mToWidth Then
    mw = mToWidth - mObj.Width
    mw = mw / Speed
    mObj.Width = mObj.Width + mw
ElseIf mObj.Width > mToWidth Then
    mw = mObj.Width - mToWidth
    mw = mw / Speed
    mObj.Width = mObj.Width - mw
End If

If mObj.Height < mToHeight Then
    mh = mToHeight - mObj.Height
    mh = mh / Speed
    mObj.Height = mObj.Height + mh
ElseIf mObj.Height > mToHeight Then
    mh = mObj.Height - mToHeight
    mh = mh / Speed
    mObj.Height = mObj.Height - mh
End If

If Round(ml) = 0 And Round(mt) = 0 And Round(mw) = 0 And Round(mh) = 0 Then
With mObj
.Left = mToLeft
.Top = mToTop
.Width = mToWidth
.Height = mToHeight
End With
cTimer.Tag = SetTag
cTimer.Enabled = False
End If
End Function

Public Function GetTaskbarHeight() As Long
    On Error Resume Next
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = (Screen.Height / Screen.TwipsPerPixelY) - rectVal.bottom
End Function

Public Sub FilterWeb(FltStr As String)
Dim sex As Boolean, bl As Boolean, cus As Boolean, Txt() As String, cut As Integer, tmps As Variant
If ReadCon("UseFilter") = 1 Then

For cut = 1 To 3
If Mid(ReadCon("Filter"), cut, 1) = 1 Then
If cut = 1 Then sex = True
If cut = 2 Then bl = True
If cut = 3 Then cus = True
End If
Next cut

If sex = True Then
tmps = Array("性交", "女优", "a片", "强奸", "露逼", "肛交", "乳交", "鸡巴", "春药", "迷奸", "换妻", "一夜情", "上床", "肏", "勃起", "吻", "裸体", "胴体", "裸聊", "手淫", "自慰", "做爱", "fuck", "porn")
For cut = 0 To UBound(tmps)
If InStr(LCase(FltStr), tmps(cut)) <> 0 Then GoTo pass
Next cut
End If

If bl = True Then
tmps = Array("爆头", "杀人", "肢解", "奸杀", "折磨", "杀害", "虐")
For cut = 0 To UBound(tmps)
If InStr(LCase(FltStr), tmps(cut)) <> 0 Then GoTo pass
Next cut
End If

If cus = True Then
Txt() = Split(ReadCon("FilterText"), ";")
For cut = 0 To UBound(Txt)
If InStr(LCase(FltStr), Txt(cut)) <> 0 Then GoTo pass
Next cut
End If

End If
Exit Sub
pass:
frmMain.picPass.Move 0, frmMain.Web(0).Top
frmMain.txtPass.SetFocus
End Sub

Sub DeleteIEHistory()
On Error Resume Next
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1", vbHide
End Sub

Sub SetBrowser(Optional Check As Boolean = False)
On Error Resume Next
Dim MeShell As String
If Check Then
If WSH.RegRead("HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\") <> "Magnifier" Then If _
MsgBox("Magnifier 目前不是您的默认浏览器，是否把 Magnifier 设为默认？", 32 + vbYesNo, "不是默认浏览器") = vbNo Then Exit Sub
End If
MeShell = MyPath & App.EXEName & ".exe"
Set WSH = CreateObject("WScript.Shell")
    WSH.RegWrite "HKEY_CLASSES_ROOT\http\shell\open\command\", """" & MeShell & """" & " " & """%1""", "REG_SZ"
    WSH.RegWrite "HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\", "Magnifier", "REG_SZ"
    WSH.RegWrite "HKEY_CLASSES_ROOT\htmlfile\shell\open\command\", """" & MeShell & """" & " " & """%1""", "REG_SZ"
    WSH.RegWrite "HKEY_CLASSES_ROOT\htmlfile\shell\open\ddeexec\Application\", "Magnifier", "REG_SZ"
    WSH.RegWrite "HKEY_CLASSES_ROOT\htmlfile\shell\opennew\command\", """" & MeShell & """" & " " & """%1""", "REG_SZ"
    WSH.RegWrite "HKEY_CLASSES_ROOT\htmlfile\shell\opennew\ddeexec\Application\", "Magnifier", "REG_SZ"
End Sub

'Public Function GetTabIndex(WebIndex As Integer) As Long
'Dim i As Long
'On Error Resume Next
'For i = 1 To frmMain.Tabs.UBound
'If frmMain.Tabs(i).Tag = WebIndex Then GetTabIndex = i: Exit For
'Next i
'End Function

Public Function CtrlInCtrl(cCtrl As Control, sCtrl As Control) As Boolean
If cCtrl.Top > sCtrl.Top And cCtrl.Top < sCtrl.Top + sCtrl.Height Then
If cCtrl.Top + cCtrl.Height > sCtrl.Top And cCtrl.Top + cCtrl.Height < sCtrl.Top + sCtrl.Height Then
CtrlInCtrl = True
End If
End If
End Function

Public Function DownloadUrlToName(Url As String) As String
On Error GoTo errH
Dim a As Integer, file As String, tmplen As Integer
For a = 1 To Len(Url)
If Left(right(Url, a), 1) = "/" Then Exit For
Next a
file = right(Url, a - 1)
If InStr(file, "?") <> 0 Then file = right(file, Len(file) - InStr(file, "?"))
If InStr(file, "_") <> 0 Then
For tmplen = -Len(file) To -1
If Mid(file, -tmplen, 1) = "." Then Exit For
Next tmplen
tmplen = -tmplen
For a = 1 To Len(file)
If a > tmplen Then If Mid(file, a, 1) = "_" Then Exit For
Next a
DownloadUrlToName = Left(file, a - 1)
End If
DownloadUrlToName = file
Exit Function
errH:
DownloadUrlToName = ""
End Function

Sub ReadCommand()
If Command = "" Then Exit Sub
frmMain.AddPage
If right(Command, 4) = ".mmb" Then
frmMain.WebGoTo LoadFile(Command)
ElseIf right(Command, 5) = ".mapx" Then
If MsgBox("您确定要安装应用 " & GetFileName(Replace(Replace(Command, """", ""), "%1", "")) _
& "吗？", 32 + vbYesNo, "安装应用") = vbYes Then
InstallApp Replace(Replace(Command, """", ""), "%1", "")
End If
Else
frmMain.WebGoTo Replace(Replace(Command, """", ""), "%1", "")
End If
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case iMsg
Case WM_GETMINMAXINFO
Dim udtMINMAXINFO As MINMAXINFO
CopyMemory udtMINMAXINFO, ByVal lParam, 40&
With udtMINMAXINFO
.ptMaxSize.X = udtMMI.ptMaxSize.X
.ptMaxSize.Y = udtMMI.ptMaxSize.Y
.ptMaxPosition.X = 0
.ptMaxPosition.Y = 0
.ptMaxTrackSize.X = .ptMaxSize.X
.ptMaxTrackSize.Y = .ptMaxSize.Y
.ptMinTrackSize.X = udtMMI.ptMinTrackSize.X
.ptMinTrackSize.Y = udtMMI.ptMinTrackSize.Y
'Debug.Print .ptMaxSize.X & "", "" & .ptMaxSize.Y
End With
CopyMemory ByVal lParam, udtMINMAXINFO, 40&
WindowProc = False
Exit Function
End Select
WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
End Function

Public Function LockWindow(hWnd As Long, Optional MinWidth As Long, Optional MinHeight As Long, Optional maxwidth As Long, Optional maxheight As Long) As Boolean
With udtMMI
If MinWidth = 0 Then .ptMinTrackSize.X = 0 Else .ptMinTrackSize.X = MinWidth
If MinHeight = 0 Then .ptMinTrackSize.Y = 0 Else .ptMinTrackSize.Y = MinHeight
If maxwidth = 0 Then .ptMaxSize.X = Screen.Width \ Screen.TwipsPerPixelX Else .ptMaxSize.X = maxwidth
If maxheight = 0 Then .ptMaxSize.Y = Screen.Width \ Screen.TwipsPerPixelX Else .ptMaxSize.Y = maxheight
End With
procOld = SetWindowLong(hWnd, -4, AddressOf WindowProc)
End Function
