Attribute VB_Name = "Mouse"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const HTCLIENT As Long = 1

Private hMouseHook As Long
Private Const KF_UP As Long = &H80000000

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Type POINTAPI
 X As Long
 Y As Long

End Type

Public Type MOUSEHOOKSTRUCT
 pt As POINTAPI
 hWnd As Long
 wHitTestCode As Long
 dwExtraInfo As Long

End Type

Public Declare Function CallNextHookEx Lib "user32" _
 (ByVal hHook As Long, _
 ByVal ncode As Long, _
 ByVal wParam As Long, _
 ByVal lParam As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" _
 Alias "SetWindowsHookExA" _
 (ByVal idHook As Long, _
 ByVal lpfn As Long, _
 ByVal hmod As Long, _
 ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" _
 (ByVal hHook As Long) As Long

Public Const WH_KEYBOARD As Long = 2
Public Const WH_MOUSE As Long = 7

Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
Public Const HC_SKIP = 2
Public Const HC_GETNEXT = 1
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE As Long = 3

Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_MOUSEWHEEL As Long = &H20A


Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const MK_RBUTTON As Long = &H2
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


Public Const VK_LBUTTON As Long = &H1
Public Const VK_RBUTTON As Long = &H2
Public Const VK_MBUTTON As Long = &H4

Dim mPt As POINTAPI
Const ptGap As Single = 5 * 5
Dim preDir As Long
Dim mouseEventDsp As String
Dim eventLength As Long

'######### mouse hook #############

Public Sub InstallMouseHook()
 hMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, _
 App.hInstance, App.ThreadID)
End Sub

Public Function MouseHookProc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Cancel As Boolean
Cancel = False
On Error GoTo due
Dim I&
Dim nMouseInfo As MOUSEHOOKSTRUCT
Dim tHWindowFromPoint As Long
Dim tpt As POINTAPI

If iCode = HC_ACTION Then
 CopyMemory nMouseInfo, ByVal lParam, Len(nMouseInfo)
 tpt = nMouseInfo.pt
 ScreenToClient nMouseInfo.hWnd, tpt
 'Debug.Print tpt.X, tpt.Y
 If nMouseInfo.wHitTestCode = 1 Then
 Select Case wParam
 Case WM_MBUTTONDOWN
 mPt = nMouseInfo.pt
 preDir = -1
 mouseEventDsp = ""
 Cancel = True
 frmMouse.Show
 frmMouse.Cls
 MouseOpen = True
 Case WM_MBUTTONUP
 frmMain.DoMouse mouseEventDsp
 Cancel = True
 Unload frmMouse
 MouseOpen = False
 Case WM_MOUSEMOVE
 If vkPress(VK_MBUTTON) Then
 Call GetMouseEvent(nMouseInfo.pt)
 End If
 End Select
 End If

End If

If Cancel Then
 MouseHookProc = 1
Else
 MouseHookProc = CallNextHookEx(hMouseHook, iCode, wParam, lParam)
End If

Exit Function

due:

End Function

Public Sub UninstallMouseHook()
 If hMouseHook <> 0 Then
 Call UnhookWindowsHookEx(hMouseHook)
 End If
 hMouseHook = 0
End Sub

Public Function vkPress(vkcode As Long) As Boolean
If (GetAsyncKeyState(vkcode) And &H8000) <> 0 Then
 vkPress = True
Else
 vkPress = False
End If
End Function

Public Function GetMouseEvent(nPt As POINTAPI) As Long
Dim cx&, cy&
Dim rtn&
rtn = -1
cx = nPt.X - mPt.X: cy = -(nPt.Y - mPt.Y)
If cx * cx + cy * cy > ptGap Then
 If cx > 0 And Abs(cy) <= cx Then
 rtn = 0
 ElseIf cy > 0 And Abs(cx) <= cy Then
 rtn = 1
 ElseIf cx < 0 And Abs(cy) <= Abs(cx) Then
 rtn = 2
 ElseIf cy < 0 And Abs(cx) <= Abs(cy) Then
 rtn = 3
 End If
 mPt = nPt
 If preDir <> rtn Then
 mouseEventDsp = mouseEventDsp & DebugDir(rtn)
 frmMouse.GetMouse mouseEventDsp
 preDir = rtn
 End If
End If
GetMouseEvent = rtn
End Function

Public Function DebugDir(nDir&) As String
Dim tStr$
Select Case nDir
 Case 0
 tStr = "ср"
Case 1
 tStr = "ио"
Case 2
 tStr = "вС"
Case 3
 tStr = "об"
Case Else
 tStr = "нч"
End Select
DebugDir = tStr
End Function

