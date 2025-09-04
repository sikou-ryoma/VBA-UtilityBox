# VBA-UtilityBox

このフォルダは、Excel VBAで業務効率化や自動化を実現するためのユーティリティ集です。  
よく使う関数やクラス、サブルーチンをモジュールごとに整理しており、ファイル選択・保存、シート・ブック管理、日付処理、ログ出力など、日々のExcel作業を支援できるツールになるよう設計しました。

ExcelのVBAプロジェクトに必要な機能だけを簡単に追加できるよう設計しています。

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
' 列番号から列記号(A, B, ..., Z, AA, ...)を取得v
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

#### SheetCollection.cls

```vba
' 全ワークシートを一括管理し、各シートのSheetManagerへアクセス
Dim coll As New SheetCollection
coll.LoadAll ThisWorkbook

Debug.Print coll.Count ' 登録されたシート数

' 全SheetManagerインスタンスのResetメソッドを一括実行
coll.ForEachCall "Reset"

' 特定シートのSheetManagerを取得してプロパティ操作
Dim sm As SheetManager
Set sm = coll.GetWs("Sheet1")
If Not sm Is Nothing Then
    sm.startRow = 2
    Debug.Print sm.startRow
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
- iniファイルのデフォルトの取り扱いについてはプロジェクトフォルダ内に`config`フォルダの作成が必要になります。
- 作成した`config`フォルダ内にファイルを置き設定を行って下さい。

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

- `Root/` ... 汎用モジュール集
- `logger/` ... ログ出力機能モジュール群
- `waitMsg/` ... 進捗時の簡易メッセージフォームと制御モジュール

## モジュール一覧

| モジュール名                | 種別      | 概要・用途                                               |
|----------------------------|-----------|----------------------------------------------------------|
| Z_ColNumToLetter.bas         | Module    | 列番号をExcelの列記号（A, B, ...）に変換する関数         |
| Z_ColumnNumber.bas          | Module    | Excel列記号に対応した列番号の列挙型定義                  |
| DateUtility.cls            | Class     | 日付操作・月初月末取得など日付ユーティリティ             |
| FileOjt.cls                | Class     | ファイル選択・保存・フォルダ選択などファイル操作支援      |
| SheetManager.cls           | Class     | シート・範囲・行列番号などワークシート管理ユーティリティ  |
| BookManager.cls            | Class     | 複数ブックの管理・切替・一括クローズ等                   |
| SheetCollection.cls        | Class     | 複数のSheetManagerインスタンスを一括管理するコレクション。キーでアクセス・追加・削除・一括操作が可能 |
| logger\clsLogger.cls       | Class     | ログ出力・ログレベル制御                                  |
| logger\Z_LogInit.bas       | Module    | ロガー初期化・設定読込                                    |
| logger\modConfig.bas       | Module    | INIファイルからの設定値取得                               |
| waitMsg\Z_WaitMsg_show.bas | Module    | 進捗メッセージフォームの表示制御                         |
| waitMsg\WaitMsg.frm        | Form      | 進捗メッセージ表示用フォーム                             |

## ライセンス

MITライセンスのもとで公開しています。詳細は `LICENSE` ファイルをご確認ください。

## 貢献

バグ報告・機能追加の提案・プルリクエスト歓迎します。

---
