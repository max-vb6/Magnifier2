VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBindStatusCallback

'����ַ����ĺ���
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
   
'���غ���
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'�������صĽӿ�
Private m_oBind As IBinding
'�Ƿ�������
Private m_fDownloading As Boolean
'�������ؿ��ƽӿڵ�������
Private m_lRefCount As Long

'���ؽ��ȵ��¼�
Public Event OnProgress(ByVal lProgress As Long, ByVal lMaxProgress As Long, ByVal lStatusCode As Long, ByVal sStatusText As String)

'��ʼ��
Private Sub Class_Initialize()
    m_fDownloading = False
    m_lRefCount = 0
End Sub

'����
Private Sub Class_Terminate()
    If m_lRefCount = 1 Then
        If Not m_oBind Is Nothing Then m_oBind.Release
    End If
    m_fDownloading = False
End Sub

'��ʼ����
Public Function StartDownloading(ByVal sSrc As String, ByVal sDest As String) As Boolean
    '����Ѿ����������˳�
    If m_fDownloading Then Exit Function
    
    Dim oBindCallback As IBindStatusCallback
    
    '���IBindStatusCallback�ӿڶ���
    Set oBindCallback = Me
    
    '��ʼ����
    DoEvents
    StartDownloading = (URLDownloadToFile(ObjPtr(Me), sSrc, sDest, 0, ObjPtr(oBindCallback)) = 0)
End Function

'��ֹ����
Public Sub AbortDownloading()
    On Error Resume Next
    If m_lRefCount = 1 Then
        If Not m_oBind Is Nothing Then m_oBind.abort
    End If
    m_fDownloading = False
End Sub

'���ַ�ָ�����ַ���
Public Function StrFromPtr(ByVal lpString As Long, Optional fUnicode As Boolean = False) As String
    On Error Resume Next
    If fUnicode Then
        StrFromPtr = String(lstrlenW(lpString), Chr(0))
        lstrcpyW StrPtr(StrFromPtr), ByVal lpString
    Else
        StrFromPtr = String(lstrlenA(lpString), Chr(0))
        lstrcpyA ByVal StrFromPtr, ByVal lpString
    End If
End Function

'*********************************************************************************************************************************************
'IBindStatusCallback�ӿڳ�Ա
'*********************************************************************************************************************************************
Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As Long, pbindinfo As Long)
'
End Sub

Private Sub IBindStatusCallback_GetPriority(pnPriority As Long)
'
End Sub

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As Long, ByVal dwSize As Long, pformatetc As Long, pstgmed As Long)
'
End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)
'
End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(ByVal riid As Long, ByVal punk As URLMonLib.IUnknownVB)
'
End Sub

'���ؽ���
Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As Long, ByVal szStatusText As Long)
    RaiseEvent OnProgress(ulProgress, ulProgressMax, ulStatusCode, StrFromPtr(szStatusText, True))
    DoEvents
End Sub

'��ʼ���ذ�
Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As URLMonLib.IBinding)
    m_fDownloading = True
    Set m_oBind = pib
    m_oBind.AddRef
    m_lRefCount = 1
End Sub

'�������ذ�
Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)
    m_fDownloading = False
    If m_lRefCount = 1 Then
        m_oBind.Release
        m_lRefCount = 0
    End If
End Sub

Private Sub IBindStatusCallback_RemoteGetBindInfo(grfBINDF As Long, pbindinfo As Long, pstgmed As Long)
'
End Sub

Private Sub IBindStatusCallback_RemoteOnDataAvailable(ByVal grfBSCF As Long, ByVal dwSize As Long, pformatetc As Long, pstgmed As Long)
'
End Sub