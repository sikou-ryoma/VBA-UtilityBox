# VBA-UtilityBox

このフォルダは、Excel VBAで業務効率化や自動化を実現するためのユーティリティ集です。  
よく使う関数やクラス、サブルーチンをモジュールごとに整理しており、  
ファイル選択・保存、シート・ブック管理、日付処理、ログ出力など、  
日々のExcel作業を支援するツールが揃っています。

ExcelのVBAプロジェクトに必要な機能だけを簡単に追加できるよう設計されています。

## 特長

- よく使うVBA関数やサブルーチンを収録
- コードの再利用性向上
- Excel作業の自動化・効率化

## インストール方法

1. 本リポジトリをダウンロードまたはクローンします。
2. ExcelのVBAエディタを開き、必要なモジュールやクラスをインポートします。

## 使い方

各モジュールや関数のコメントを参照してください。サンプルコードも随時追加予定です。

## 使用例

#### ColNumToLetter.bas

```vba
' 列番号から列記号(A, B, ..., Z, AA, ...)を取得
Dim colLetter As String
colLetter = ColumnNumberToLetter(28) ' → "AB"
```

#### DateUtility.cls

```vba
Dim du As New DateUtility
du.targetYear = 2024
du.targetMonth = 6
du.targetDay = 1
Debug.Print du.FirstDayOfMonth   ' → 2024/6/1
Debug.Print du.LastDayOfMonth    ' → 2024/6/30
```

#### FileOjt.cls

```vba
Dim fo As New FileOjt
Dim filePath As Variant
filePath = fo.OpFile() ' ファイル選択ダイアログ表示
If filePath <> False Then
    MsgBox "選択ファイル: " & filePath
End If
```

#### SheetManager.cls

```vba
Dim sm As New SheetManager
Set sm.sheet = ThisWorkbook.Sheets(1)
Debug.Print sm.lastRow() ' シートの最終行番号を取得
```

#### BookManager.cls

```vba
Dim bm As New BookManager
bm.AddWorkbook "main", "C:\test.xlsx"
Dim wb As Workbook
Set wb = bm.GetWorkbook("main")
If Not wb Is Nothing Then
    MsgBox wb.Name
End If
```

#### logger（ログ出力機能）

```vba
' 初期化（パスは任意のログ保存先フォルダを指定）
Call InitializeLogger(ThisWorkbook.path)

' ログ出力
logger.Info "処理開始"
logger.DebugMsg "デバッグ情報"
logger.WarnMsg "警告メッセージ"
logger.ErrorMsg "エラー発生"
```

#### waitMsg（進捗メッセージフォーム）

```vba
' 指定秒数だけ進捗メッセージフォームを表示
Call waitMsgShow(3)  ' 3秒間表示

' 何らかの処理

Unload WaitMsg       ' 表示終了
```

'テスト用サンプル
```vba
Sub test()
    Call waitMsgShow
    ' ここに処理を書く
    Unload WaitMsg
End Sub
```

## フォルダ構成

- `Modules/` ... 汎用モジュール
- `Classes/` ... クラスモジュール
- `Samples/` ... サンプルファイル

## ライセンス

MITライセンスのもとで公開しています。詳細は `LICENSE` ファイルをご確認ください。

## 貢献

バグ報告・機能追加の提案・プルリクエスト歓迎します。

---
