Attribute VB_Name = "Theme"
Public Type Theme
MainColor As Long
ToolBar As Long
TabOnAct As Long
TabNoAct As Long
End Type

Public ThmClr As Theme

Public Sub LoadTheme()
On Error GoTo ThmErr
Dim ClrFile As String
ClrFile = MyPath & "Themes\" & ReadThm("Theme")
With ThmClr
.MainColor = ReadString("Clr", "Main", ClrFile)
.TabOnAct = ReadString("Clr", "TabAct", ClrFile)
.TabNoAct = ReadString("Clr", "TabNrm", ClrFile)
.ToolBar = ReadString("Clr", "ToolBar", ClrFile)
End With
With frmMain
.BackColor = ThmClr.ToolBar
.picCaption.BackColor = ThmClr.ToolBar
.picTabBar.BackColor = ThmClr.ToolBar
.picAddr.BackColor = ThmClr.ToolBar
.lstAddr.BackColor = ThmClr.MainColor
.lstAddr.SelColor = ThmClr.ToolBar
.picSta.BackColor = ThmClr.ToolBar
.txtPass.BackColor = ThmClr.ToolBar
.txtAddr.BackColor = ThmClr.MainColor
.picZoom.BackColor = ThmClr.MainColor
.picPass.BackColor = ThmClr.MainColor
.picPassCore.BackColor = .picPass.BackColor
.imgBtn(0).Picture = LoadPicture(GetThmFolder & "Lf.bmp")
.imgBtn(1).Picture = LoadPicture(GetThmFolder & "Rf.bmp")
.imgBtn(2).Picture = LoadPicture(GetThmFolder & "St.bmp")
.imgBtn(3).Picture = LoadPicture(GetThmFolder & "Wr.bmp")
.imgCtrl(0).Picture = LoadPicture(GetThmFolder & "Cl.bmp")
.imgCtrl(1).Picture = LoadPicture(GetThmFolder & "Mx.bmp")
.imgCtrl(2).Picture = LoadPicture(GetThmFolder & "Mn.bmp")
.imgCtrl(3).Picture = LoadPicture(GetThmFolder & "Ap.bmp")
.picAddr.Picture = LoadPicture(GetThmFolder & "Tx.bmp")
For i = 0 To 3
.imgBtn(i).Move .imgBtn(i).Left, (.picCaption.Height - .imgBtn(i).Height) / 2
.imgCtrl(i).Move .imgCtrl(i).Left, (.picCaption.Height - .imgCtrl(i).Height) / 2
Next i
If ReadString("Clr", "WhiteFont", ClrFile) = 1 Then
.lblShow(0).ForeColor = vbWhite
.lblShow(1).ForeColor = vbWhite
.lblZmVl.ForeColor = vbWhite
.lblPlMi(0).ForeColor = vbWhite
.lblPlMi(1).ForeColor = vbWhite
.lblSta.ForeColor = vbWhite
.txtAddr.ForeColor = vbWhite
.txtPass.ForeColor = vbWhite
.lstAddr.ForeColor = vbWhite
End If
End With
Exit Sub
ThmErr:
MsgBox "浏览器在加载主题时发生了一个错误，现在被迫关闭", 48, "非常抱歉"
End
End Sub

Public Function GetThmFolder() As String
On Error Resume Next
GetThmFolder = MyPath & "Themes\" & GetFileName(ReadThm("Theme")) & "\"
End Function
