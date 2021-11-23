Attribute VB_Name = "Apps"
Public Function RunApp(AppIniFile As String) As String
On Error GoTo AppErr
Dim sName As String, sScript As String, sType As String, iUseE As Integer
sName = ReadString("App", "Name", AppIniFile)
If LCase(ReadString("App", "Type", AppIniFile)) = "vbscript" Then
sType = "VBScript"
ElseIf LCase(ReadString("App", "Type", AppIniFile)) = "javascript" Then
sType = "JScript"
Else
frmMain.WebGoTo LoadFile(MyPath & ReadString("App", "Url", AppIniFile))
Exit Function
End If
sScript = LoadFile(MyPath & ReadString("App", "Script", AppIniFile))
If sScript <> "" Then
Load frmSHide.SC(frmSHide.SC.UBound + 1)
With frmSHide.SC(frmSHide.SC.UBound)
.language = sType
.Tag = sName
.Reset
If ReadString("App", "UseExtent", AppIniFile) = 1 Then
If ReadString("App", "AllowedExt", AppIniFile) = 0 Then
If MsgBox("应用 " & sName & _
" 请求控制您的网页。确认运行该应用吗？" & vbCrLf & _
"（您的选择将会被记住，以后将不再询问）", vbYesNo + 48, "请求权限") = vbNo Then Exit Function Else WriteString "App", "AllowedExt", 1, AppIniFile
End If
.AddObject "Url", URL, True
End If
If ReadString("App", "HasUI", AppIniFile) = 1 Then
Dim AppForm As New frmApp
AppForm.Show
AppForm.Tag = frmSHide.SC.UBound
AppForm.SetHtml MyPath & ReadString("App", "UI", AppIniFile)
.AddObject "Form", AppForm, True
.AddObject "Html", AppForm.WbHtml, True
End If
Dim cApp As New AppControl
cApp.scIndex = frmSHide.SC.UBound
.AddObject "AppControl", cApp, True
DoEvents
.AddCode sScript
Set cApp = Nothing
End With
End If
Exit Function
AppErr:
RunApp = "应用 " & sName & " 运行时出现错误" & vbCrLf & "错误信息: " & Err.Description & vbCrLf & _
"行 " & frmSHide.SC(frmSHide.SC.UBound).Error.Line & " 列 " & frmSHide.SC(frmSHide.SC.UBound).Error.Column
End Function

Public Function ReadAppAbout(AppIndex As Long) As String
On Error Resume Next
ReadAppAbout = ReadString("App", "About", MyPath & "Apps\" & ReadApp("app" & AppIndex))
End Function


Sub LoadApps()
On Error Resume Next
With frmMain
If ReadApp("Count") = 0 Then .picAppCore.Visible = False: .lblShow(2).Visible = True: Exit Sub
Dim I As Long, sIco As String, lWth As Long, lx As Long, ly As Long, lcx As Long, lcy As Long

If .picAIco.UBound > 0 Then
For I = 1 To .picAIco.UBound
Unload .picAIco(I)
Next I
End If                            'Delete All Icon

For I = 1 To ReadApp("Count")
Load .picAIco(I)

sIco = ReadString("App", "Icon", MyPath & "Apps\" & ReadApp("app" & I))
If sIco <> "" Then .picAIco(I).Picture = LoadPicture(MyPath & sIco)

lWth = 48 * Screen.TwipsPerPixelX
lx = (.picApp.Width - lWth * 5) / 6
ly = (.picApp.Height - lWth * 3) / 4

lcx = I Mod 5
If lcx = 0 Then lcx = 5
lcy = Int(I / 5) + 1
If I Mod 5 = 0 Then lcy = lcy - 1

.picAIco(I).Move lcx * lx + (lcx - 1) * lWth, lcy * ly + (lcy - 1) * lWth, lWth, lWth
.picAIco(I).ToolTipText = ReadString("App", "Name", MyPath & "Apps\" & ReadApp("app" & I)) & "    版权: " & ReadAppAbout(I)
.picAIco(I).Visible = True
Next I

.picAppCore.Visible = True
.picAppCore.Height = .picAIco(.picAIco.UBound).Top + .picAIco(.picAIco.UBound).Height + ly
.lblShow(2).Visible = False
If .picAppCore.Height > .picApp.Height Then
.sroApp.Max = .picAppCore.Height - .picApp.Height
.sroApp.SmallChange = lcy + lWth
.sroApp.LargeChange = (lcy + lWth) * 2
.sroApp.Visible = True
Else
.sroApp.Visible = False
End If
End With
End Sub

Sub DeleteApp(AppIndex As Integer)
On Error Resume Next
Dim I As Long, ToI As Long, AppsStr As String, Apps() As String, FileStr As String, Fldr As String
ToI = ReadApp("Count")

Fldr = ReadApp("app" & AppIndex)
For I = 1 To Len(Fldr)
If Mid(Fldr, I, 1) = "\" Then Exit For
Next I
Fldr = MyPath & "Apps\" & Left(Fldr, I - 1)
DeleteFolder Fldr

For I = 1 To ToI
If I = AppIndex And I = ToI Then Exit For
If I = AppIndex Then I = I + 1
AppsStr = AppsStr & ReadApp("app" & I) & "@@"
Next I

Apps = Split(AppsStr, "@@")
FileStr = "[Apps]" & vbCrLf & "Count=" & ToI - 1

For I = 0 To UBound(Apps) - 1
FileStr = FileStr & vbCrLf & "app" & I + 1 & "=" & Apps(I)
Next I

Open MyPath & "Apps\Apps.ini" For Output As #2
Print #2, FileStr
Close #2

LoadApps
End Sub

Sub InstallApp(MapxFile As String)
On Error Resume Next
If Dir(MyPath & "rar.exe") = "" Then MsgBox "应用解压缩模块已被删除，无法正常安装应用", 48, "应用安装时错误": Exit Sub

Dim Orgn As String, ApNm As String, Cnt As Long
ApNm = GetFileName(MapxFile)

Clipboard.SetText MyPath & "rar.exe x """ & MapxFile & """ """ & MyPath & "Apps\" & ApNm & """"
Shell MyPath & "rar.exe x """ & MapxFile & """ """ & MyPath & "Apps\""", vbHide

ApNm = ApNm & "\" & ApNm & ".ini"
Cnt = ReadApp("Count")

If Cnt > 0 Then
For I = 1 To Cnt
If ReadApp("app" & I) = ApNm Then
GoTo Over
Exit For
End If
Next I
End If

SaveApp "Count", Cnt + 1
Orgn = LoadFile(MyPath & "Apps\Apps.ini")
Orgn = Orgn & "app" & Cnt + 1 & "=" & ApNm

Open MyPath & "Apps\Apps.ini" For Output As #3
Print #3, Orgn
Close #3

DoEvents: Sleep 100
WriteString "App", "AllowedExt", 0, MyPath & "Apps\" & ApNm

Over:
frmMain.tmrSApp.Enabled = True
End Sub

Sub RunOnLoadApp()
On Error Resume Next
Dim I As Long
For I = 1 To ReadApp("Count")
If ReadString("App", "OnLoad", MyPath & "Apps\" & ReadApp("app" & I)) = 1 Then RunApp MyPath & "Apps\" & ReadApp("app" & I)
Next I
End Sub

Public Function SearchApp(SrchStr As String) As String
On Error Resume Next
Dim I As Long, Rslts As String
For I = 1 To CLng(ReadApp("Count"))
If InStr(LCase(ReadString("App", "Name", MyPath & "Apps\" & ReadApp("app" & I))), LCase(SrchStr)) <> 0 Then
Rslts = Rslts & ReadString("App", "Name", MyPath & "Apps\" & ReadApp("app" & I)) & "@@"
End If
Next I
SearchApp = Rslts
End Function
