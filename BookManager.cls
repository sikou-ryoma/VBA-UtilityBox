VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASSNAME As String = "BookManager"

'ブックの複数管理用ユーティリティクラス
'Dictionaryでの管理

Private wbDict As Object

'ライブラリの呼び出しはインスタンス生成時に自動で行う
Private Sub Class_Initialize()
    Set wbDict = CreateObject("Scripting.Dictionary")
End Sub

'ブックのディクショナリへの追加
'ディクショナリを使用するためaddにはkeyとフルパスが必要
'IsWorkbookOpenを呼び出して指定したブックが開いているかを判断し適切に処理
Public Sub AddWorkbook(ByVal key As String, ByVal filePath As String)
                                                    
                                                    
    Dim wbName As String
    wbName = Dir(filePath)
    
    Dim k As Variant
    
    For Each k In wbDict.keys
        If wbDict(k).Name = wbName Then
            MsgBox "ディクショナリ内に同じ名前のブックが存在します。", vbExclamation, CLASSNAME
            Exit Sub
        End If
    Next k
    
    If Not wbDict.exists(key) Then
        Dim wb As Workbook
        If IsWorkbookOpen(wbName) Then
            Set wb = Workbooks(wbName)
            wbDict.Add key, wb
        Else
            Set wb = Workbooks.Open(filePath)
            wbDict.Add key, wb
        End If
    Else
        MsgBox "キー '" & key & "' は既に存在します。", vbExclamation, CLASSNAME
    End If
    
End Sub

'ディクショナリよりブックの呼び出し
Public Function GetWorkbook(ByVal key As String) As Workbook

    If wbDict.exists(key) Then
        Set GetWorkbook = wbDict(key)
    Else
        Set GetWorkbook = Nothing
    End If
    
End Function

'ディクショナリの確認用
Public Sub PrintWorkbookNames()

    Dim k As Variant
    
    For Each k In wbDict.keys
        Debug.Print k & ": " & wbDict(k).Name
    Next k
    
End Sub

'指定したブックをディクショナリから削除
Public Sub CloseAndRemove(ByVal key As String, Optional ByVal saveChanges As Boolean = False)

    If wbDict.exists(key) Then
        wbDict(key).Close saveChanges:=saveChanges
        wbDict.Remove key
    End If
    
End Sub

'全てのブックをディクショナリから削除
Public Sub CloseAll(Optional ByVal saveChanges As Boolean = False)

    Dim k As Variant
    
    For Each k In wbDict.keys
        wbDict(k).Close saveChanges:=saveChanges
    Next k
    
    wbDict.RemoveAll
    
End Sub

'指定したブックが開いているかを確認する関数
'AddWorkbook用
Private Function IsWorkbookOpen(ByVal wbName As String) As Boolean

    Dim wb As Workbook
    
    On Error Resume Next
    Set wb = Workbooks(wbName)
    IsWorkbookOpen = Not wb Is Nothing
    On Error GoTo 0
    
End Function


