VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsPathConfig


Option Explicit

Private mProjectRoot As String
Private mReportsPath As String
Private mTempPath As String
Private mCustomPaths As Object ' 可変パスを保持する辞書

Private Sub Class_Initialize()
    '---可変パス用のDictionary
    Set mCustomPaths = CreateObject("Scripting.Dictionary")
End Sub

'---固定系パス
Public Property Get ProjectRoot() As String
    ProjectRoot = mProjectRoot
End Property

Public Property Get ReportsPath() As String
    ReportsPath = mReportsPath
End Property

Public Property Get TempPath() As String
    TempPath = mTempPath
End Property


'---固定パスのイニシャライズ
'---Initの呼び出しは必ず行う
Public Sub Init(ByVal basePath As String)
    mProjectRoot = basePath
    mReportsPath = mProjectRoot & "\output\reports"
    mTempPath = mProjectRoot & "\output\temp"
End Sub


'---カスタムパス
Public Sub SetPath(ByVal key As String, ByVal path As String)
    If mCustomPaths.Exists(key) Then
        mCustomPaths(key) = path
    Else
        mCustomPaths.Add key, path
    End If
End Sub

Public Function GetPath(ByVal key As String) As String
    If mCustomPaths.Exists(key) Then
        GetPath = mCustomPaths(key)
    Else
        Err.Raise 9999, , "Path for key '" & key & "' is not set."
    End If
End Function

