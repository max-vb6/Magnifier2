Attribute VB_Name = "Main"
Option Explicit

Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const MAX_PATH = 260
Public Const ILD_TRANSPARENT = &H1
Public Type SHFILEINFO
hIcon As Long
iIcon As Long
dwAttributes As Long
szDisplayName As String * MAX_PATH
szTypeName As String * 80
End Type
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long
Public shinfo As SHFILEINFO
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const SHGFI_ICON = &H100

Public Function GetFileIcon(fName As String, sPicture As PictureBox)
Dim r As Long, hImgLarge As Long
hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_LARGEICON Or BASIC_SHGFI_FLAGS Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
sPicture.Picture = LoadPicture()
sPicture.AutoRedraw = True
r = ImageList_Draw(hImgLarge&, shinfo.iIcon, sPicture.hDC, 0, 0, ILD_TRANSPARENT)
Set sPicture.Picture = sPicture.Image
GetFileIcon = r
End Function

Public Function GetFileName(Path As String, Optional GetEx As Boolean) As String
On Error GoTo FileErr
Dim tstrs() As String
tstrs = Split(Path, "\")
If GetEx Then GetFileName = tstrs(UBound(tstrs)): Exit Function
tstrs = Split(tstrs(UBound(tstrs)), ".")
GetFileName = tstrs(0)
Exit Function
FileErr:
GetFileName = ""
End Function

