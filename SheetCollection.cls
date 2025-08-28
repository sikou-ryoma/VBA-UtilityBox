VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-----------------------------
' 複数の SheetManager インスタンスを一括管理するためのクラス
' 各ワークシートごとに既存の SheetManager クラスを使ってインスタンスを生成
' 名前（キー）でアクセス、追加、削除が可能
' 全体に対するメソッド呼び出し（CallByName）をサポート
'
' 前提: プロジェクトにクラスモジュール "SheetManager" が存在すること
'-----------------------------

Option Explicit

Private pDict As Object ' Scripting.Dictionary

'---クラス初期化
Private Sub Class_Initialize()
    Set pDict = CreateObject("Scripting.Dictionary")
End Sub

'---クラス解放
Private Sub Class_Terminate()
    On Error Resume Next
    Set pDict = Nothing
    On Error GoTo 0
End Sub

'--------------------------------------------------
' AddWs - ワークシートから新しい SheetManager を生成して追加
'   ws: Worksheet オブジェクト
'   Optional key: コレクションに登録するキー（省略すると ws.Name を使用）
'   既に同じキーがある場合はエラーを発生させる
'--------------------------------------------------
Public Sub AddWs(ByVal ws As Worksheet, Optional ByVal key As String)
    Dim sm As SheetManager
    If ws Is Nothing Then Err.Raise vbObjectError + 1000, "SheetCollection.Add", "Worksheet が無効です"

    If key = "" Then key = ws.Name
    If pDict.Exists(key) Then Err.Raise vbObjectError + 1001, "SheetCollection.Add", "キー '" & key & "' は既に存在します"

    Set sm = New SheetManager
    Set sm.sheet = ws
    pDict.Add key, sm
End Sub

'--------------------------------------------------
' Remove - キーで削除
'--------------------------------------------------
Public Sub Remove(ByVal key As String)
    If pDict.Exists(key) Then pDict.Remove key
End Sub

'--------------------------------------------------
' Get - キーから SheetManager を返す（存在しなければ Nothing）
'--------------------------------------------------
Public Function GetWs(ByVal key As String) As SheetManager
    If pDict.Exists(key) Then
        Set GetWs = pDict(key)
    Else
        Set GetWs = Nothing
    End If
End Function

'--------------------------------------------------
' GetByIndex - 1-based インデックスで取得（For i = 1 To Count の想定）
'--------------------------------------------------
Public Function GetByIndex(ByVal idx As Long) As SheetManager
    If idx < 1 Or idx > pDict.Count Then
        Set GetByIndex = Nothing
    Else
        Dim k As Variant
        k = pDict.Keys()(idx - 1) ' Keys は 0-based の配列を返す
        Set GetByIndex = pDict(k)
    End If
End Function

'--------------------------------------------------
' Count - 登録数
'--------------------------------------------------
Public Property Get Count() As Long
    Count = pDict.Count
End Property

'--------------------------------------------------
' Keys - 登録されているキーの配列を返す
'--------------------------------------------------
Public Property Get Keys() As Variant
    If pDict.Count = 0 Then
        Keys = Array()
    Else
        Keys = pDict.Keys
    End If
End Property

'--------------------------------------------------
' Items - 登録されている SheetManager インスタンスを Collection として返す
'        （VBA 側で For Each を使いたいときに便利）
'--------------------------------------------------
Public Function Items() As Collection
    Dim coll As New Collection
    Dim key As Variant
    For Each key In pDict.Keys
        coll.Add pDict(key)
    Next key
    Set Items = coll
End Function

'--------------------------------------------------
' LoadAll - 指定ブック内の全ワークシートを読み込み（既存キーはスキップ）
'   Optional wb: 読み込む Workbook（省略で ThisWorkbook）
'--------------------------------------------------
Public Sub LoadAll(Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        On Error Resume Next
        If Not pDict.Exists(ws.Name) Then Add ws, ws.Name
        On Error GoTo 0
    Next ws
End Sub

'--------------------------------------------------
' ForEachCall - 各 SheetManager に対してメソッドを一括実行
'   methodName: SheetManager 側の公開メソッド名
'   ParamArray args: メソッド引数（必要な場合）
'   例: coll.ForEachCall "Reset"
'        coll.ForEachCall "SetStartRow", 2
'--------------------------------------------------
Public Sub ForEachCall(ByVal methodName As String, ParamArray args() As Variant)
    Dim key As Variant
    Dim obj As Object
    For Each key In pDict.Keys
        Set obj = pDict(key)
        If Not obj Is Nothing Then
            ' CallByName でメソッド呼び出し
            On Error Resume Next
            Select Case UBound(args)
                Case -1
                    CallByName obj, methodName, VbMethod
                Case 0
                    CallByName obj, methodName, VbMethod, args(0)
                Case 1
                    CallByName obj, methodName, VbMethod, args(0), args(1)
                Case 2
                    CallByName obj, methodName, VbMethod, args(0), args(1), args(2)
                Case 3
                    CallByName obj, methodName, VbMethod, args(0), args(1), args(2), args(3)
                Case Else
                    ' 4 つ以上の引数が必要な場合は適宜拡張してください
                    CallByName obj, methodName, VbMethod, args
            End Select
            If Err.Number <> 0 Then
                ' エラーは無視して次へ（必要ならログ出力する）
                Debug.Print "ForEachCall error on key=" & key & " method=" & methodName & " Err=" & Err.Number & ":" & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next key
End Sub

'--------------------------------------------------
' Exists - キーの存在確認
'--------------------------------------------------
Public Function Exists(ByVal key As String) As Boolean
    Exists = pDict.Exists(key)
End Function

'--------------------------------------------------
' Clear - すべての登録を破棄（各インスタンスの解放も）
'--------------------------------------------------
Public Sub Clear()
    Dim key As Variant
    For Each key In pDict.Keys
        On Error Resume Next
        Set pDict(key) = Nothing
        On Error GoTo 0
    Next key
    pDict.RemoveAll
End Sub

'-----------------------------
' 使い方の例（標準モジュールにコピーして実行）
'-----------------------------
' Sub Example_UseSheetCollection()
'     Dim coll As New SheetCollection
'     ' ThisWorkbook の全ワークシートを読み込む
'     coll.LoadAll ThisWorkbook
'
'     Debug.Print "登録数: " & coll.Count
'
'     ' 全インスタンスの Reset メソッドを呼ぶ（SheetManager に Reset が実装されている前提）
'     coll.ForEachCall "Reset"
'
'     ' 特定シートの SheetManager を取得してプロパティを操作
'     Dim sm As SheetManager
'     Set sm = coll.Get("Sheet1")
'     If Not sm Is Nothing Then
'         ' 例: sm.StartRow = 2
'         '     sm.SomeMethod など
'     End If
' End Sub

'-----------------------------
' 注意点
' - SheetManager クラス内のメンバー名（プロパティやメソッド）に合わせて ForEachCall や Get の使い方を調整してください。
' - CallByName を使うため、対象となるメソッドは Public である必要があります。
' - 4 つ以上の引数を渡す ForEachCall の需要がある場合は Select Case 部分を拡張してください。
'-----------------------------

