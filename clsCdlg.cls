VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCdlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pSavefilename As OPENFILENAME) As Long
Private Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type
Public FileName As String
Public rtnFilter As Long

Public Function ShowOpen(hwnd As Long, Filter As String, Title As String) As String
Dim OFName As OPENFILENAME
With OFName
.lStructSize = Len(OFName)
.hwndOwner = hwnd
.hInstance = App.hInstance
.lpstrFilter = Filter
.lpstrFile = Space$(259)
.nMaxFile = 260
.lpstrFileTitle = Space$(259)
.nMaxFileTitle = 260
.lpstrInitialDir = MyPath
.lpstrTitle = Title
.flags = 0
End With
If GetOpenFileName(OFName) Then
ShowOpen = Trim(OFName.lpstrFile)
Else
ShowOpen = ""
End If
FileName = Replace(ShowOpen, Chr(0), "")
End Function

Public Function ShowSave(hwnd As Long, Filter As String, Title As String) As String
Dim OFName As OPENFILENAME
With OFName
.lStructSize = Len(OFName)
.hwndOwner = hwnd
.hInstance = App.hInstance
.lpstrFilter = Filter
.lpstrFile = FileName & String(255 - Len(FileName), vbNullChar)
.nMaxFile = 260
.lpstrFileTitle = Space$(259)
.nMaxFileTitle = 260
.lpstrInitialDir = MyPath
.lpstrTitle = Title
.flags = 0
End With
If GetSaveFileName(OFName) Then
ShowSave = Trim(OFName.lpstrFile)
Else
ShowSave = ""
End If
FileName = Replace(ShowSave, Chr(0), "")
rtnFilter = OFName.nFilterIndex
End Function
