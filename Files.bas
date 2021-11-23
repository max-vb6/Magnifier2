Attribute VB_Name = "Files"
Public Enum FO_Operation
FO_MOVE = 1
FO_COPY = 2
FO_DELETE = 3
FO_RENAME = 4
End Enum

Public Enum FOFlags
FOF_MULTIDESTFILES = &H1 'Destination specifies multiple files
FOF_SILENT = &H4 'Don't display progress dialog
FOF_RENAMEONCOLLISION = &H8 'Rename if destination already exists
FOF_NOCONFIRMATION = &H10 'Don't prompt user
FOF_WANTMAPPINGHANDLE = &H20 'Fill in hNameMappings member
FOF_ALLOWUNDO = &H40 'Store undo information if possible
FOF_FILESONLY = &H80 'On *.*, don't copy directories
FOF_SIMPLEPROGRESS = &H100 'Don't show name of each file
FOF_NOCONFIRMMKDIR = &H200 'Don't confirm making any needed dirs
End Enum

Public Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private op As SHFILEOPSTRUCT

'===Config===
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'====End=====

Public Sub DeleteFolder(sDeleteFolder As String, Optional Interface As Boolean = False)
SetAttr sDeleteFolder, vbNormal
With op
.wFunc = FO_DELETE
.pFrom = sDeleteFolder
.fFlags = IIf(Interface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
End With
SHFileOperation op
End Sub

Public Function LoadFile(FilePath As String) As String
On Error GoTo FileErr
Dim sFl As String, iFreeNum As Integer
iFreeNum = FreeFile
Open FilePath For Binary As #iFreeNum
sFl = Space(LOF(iFreeNum))
Get #iFreeNum, , sFl
Close #iFreeNum
LoadFile = sFl
Exit Function
FileErr:
LoadFile = ""
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

Public Function FileCopyEx(ByVal SouFileName As String, ByVal DestFileName As String)
'复制文件,可以复制正在使用的文件.
'SouFileName - 源文件
'DestFileName - 目标文件
'By 嗷嗷叫的老马
'2007-05-05
Dim tmpArr() As Byte
Open SouFileName For Binary Access Read As #1
ReDim tmpArr(LOF(1))
Get 1, , tmpArr
Close #1
Open DestFileName For Binary As #2
Put 2, , tmpArr
Close #2
ReDim tmpArr(0)             '释放内存
End Function

Public Function SaveHistory(URL As String) As Long
On Error Resume Next
Dim HisDat As String, SavDat As String, I As Long
HisDat = LoadFile(MyPath & "Data\InputHis.ini")
If HisDat = "" Then
SavDat = Trim(URL) & "@@"
Else
HisDat = Replace(HisDat, Trim(URL) & "@@", "")
SavDat = Trim(URL) & "@@" & HisDat
End If
Open MyPath & "Data\InputHis.ini" For Output As #4
Print #4, SavDat
Close #4
End Function

Public Function SearchHistory(SrchStr As String) As String
On Error Resume Next
Dim HisLst() As String, I As Long, Rslts As String
HisLst = Split(LoadFile(MyPath & "Data\InputHis.ini"), "@@")
For I = 0 To UBound(HisLst) - 1
If InStr(HisLst(I), SrchStr) <> 0 Then Rslts = Rslts & HisLst(I) & "@@"
Next I
SearchHistory = Rslts
End Function

Public Sub DeleteHistory()
On Error Resume Next
Open MyPath & "Data\InputHis.ini" For Output As #5
Print #5, ""
Close #5
End Sub

Public Function SaveBookmark(sUrl As String, sCaption As String) As Long
On Error Resume Next
Open MyPath & "Data\Bookmarks\" & sCaption & ".mmb" For Output As #6
Print #6, sUrl
Close #6
End Function

'===Config===
Public Function ReadString(ByVal Caption As String, ByVal item As String, ByVal Path As String) As String
    On Error Resume Next
    Dim sBuffer As String
    sBuffer = Space(128)
    GetPrivateProfileString Caption, item, vbNullString, sBuffer, 128, Path
    
    ReadString = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function

Public Function MyPath() As String
    Dim sPath As String
    sPath = App.Path
    
    If right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    MyPath = sPath
End Function

Public Function WriteString(ByVal Caption As String, ByVal item As String, ByVal ItemValue As String, ByVal Path As String) As Long
    Dim sBuffer As String
    sBuffer = Space(128)
    
    sBuffer = ItemValue & vbNullChar
    WriteString = WritePrivateProfileString(Caption, item, sBuffer, Path)
End Function

Public Function SaveCon(item As String, Txt As String) As Long
WriteString "Setting", item, Txt, MyPath & "config.ini"
End Function

Public Function ReadCon(item As String) As String
ReadCon = ReadString("Setting", item, MyPath & "config.ini")
End Function

Public Function SaveThm(item As String, Txt As String) As String
WriteString "Theme", item, Txt, MyPath & "config.ini"
End Function

Public Function ReadThm(item As String) As String
ReadThm = ReadString("Theme", item, MyPath & "config.ini")
End Function

Public Function SaveApp(item As String, Txt As String) As String
WriteString "Apps", item, Txt, MyPath & "Apps\Apps.ini"
End Function

Public Function ReadApp(item As String) As String
ReadApp = ReadString("Apps", item, MyPath & "Apps\Apps.ini")
End Function
'====End=====
