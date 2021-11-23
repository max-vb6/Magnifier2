VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public scIndex As Long

Sub AppSleep(ByVal dwMilliseconds As Long)
    Dim SaveTime As Long
    Dim NowTime As Long
    Dim IsWait As Long
    IsWait = 0
    SaveTime = GetTickCount
    Do
       DoEvents
       NowTime = GetTickCount
       If NowTime - SaveTime >= dwMilliseconds Then
          IsWait = 1
       End If
    Loop While IsWait = 0
End Sub

Sub GoToUrl(sUrl As String)
frmMain.WebGoTo sUrl
End Sub

Sub AddTab()
frmMain.AddPage
End Sub

Function AppPath() As String
AppPath = MyPath
End Function

Function GetFile(sPath As String) As String
GetFile = LoadFile(sPath)
End Function

Sub OpenPassword()
frmMain.picPass.Top = frmMain.picCaption.Height
frmMain.txtPass.SetFocus
End Sub

Sub DelBrwHistory()
DeleteIEHistory
End Sub

Sub DelInuptHistory()
DeleteHistory
End Sub

Sub NewDownload(sUrl As String, sPath As String)
If Dir(MyPath & "MagDown.exe") <> "" Then
Shell MyPath & "MagDown.exe " & sUrl & "@@" & sPath, vbNormalFocus
End If
End Sub

Sub EndApp()
On Error Resume Next
frmSHide.SC(scIndex).Reset
Unload frmSHide.SC(scIndex)
End Sub
