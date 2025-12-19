# Sheet6 UIセットアップガイド

## 概要

Sheet6.clsにはフォルダパス入力機能が実装されました。Excelのワークシート上に以下の要素を配置してください。

## 実装済みの機能

- フォルダ選択ボタンのクリック処理: `onClick_SelectFolder()`
- 実行ボタンのクリック処理: `onClick_Exec_Button()`
- フォルダパス入力欄: セル `B3` に自動入力
- Sheet2~5の関数の自動実行

## セットアップ手順

### 1. テキストボックスの追加

Sheet6に以下を配置してください：

**ラベル (A3)**

- 内容: "フォルダパス："
- 位置: セルA3

**テキストボックス (B3)**

- 位置: セルB3に合わせて配置
- 名前: フォルダパスを表示する入力欄として使用
- 注: VBAコードで自動的に入力される

### 2. ボタンの追加

**フォルダ選択ボタン**

- 配置位置: セルC3またはD3
- ボタンテキスト: "フォルダ選択"
- マクロ割当: `onClick_SelectFolder`
- 操作方法:
  1. Excelで「挿入」→「ボタン（フォームコントロール）」を選択
  2. Sheet6上でボタンのサイズを決めて配置
  3. 右クリック → 「マクロの登録」
  4. 登録するマクロ: `Sheet6.onClick_SelectFolder`

**実行ボタン**

- 配置位置: セルC4またはD4
- ボタンテキスト: "実行"
- マクロ割当: `onClick_Exec_Button`
- 操作方法: フォルダ選択ボタンと同様に設定
  1. 「挿入」→「ボタン（フォームコントロール）」を選択
  2. Sheet6上でボタンを配置
  3. 右クリック → 「マクロの登録」
  4. 登録するマクロ: `Sheet6.onClick_Exec_Button`

### 3. 使用方法

1. **Sheet6にアクセス**: ExcelでSheet6を開く
2. **フォルダ選択**: 「フォルダ選択」ボタンをクリック
   - フォルダ選択ダイアログが表示される
   - フォルダを選択すると、B3にパスが自動入力される
3. **実行**: 「実行」ボタンをクリック
   - Sheet2~5の関数が順序に実行される
   - 各シートの結合処理が実行される
   - 完了メッセージが表示される

## VBA関数の説明

### `onClick_SelectFolder()`

- フォルダ選択ダイアログを開く
- 選択されたフォルダパスを `Range("B3")` に入力

### `onClick_Exec_Button()`

- `Range("B3")` のフォルダパスを取得
- 入力チェック（空かどうか、フォルダが存在するか）
- Sheet2.sheet_1_2()、Sheet3.sheet_1_3() など4つの関数を順序実行
- 完了メッセージを表示

### `select_folder()`

- プライベート関数
- `Application.FileDialog()` を使用してフォルダ選択ダイアログを表示
- 選択されたフォルダパスを戻す

## Module1の更新内容

`Concat_Sheet()` 関数に `folderPath` パラメータが追加されました：

```vba
Public Function Concat_Sheet( _
    ByVal destSheetName As String, _
    ByVal srcSheetName As String, _
    ByVal destStartRow As Long, _
    ByVal srcStartRow As Long, _
    ByVal copyColCount As Long, _
    ByVal folderPath As String _
) As Long
```

- Sheet2~5から呼び出される際に、フォルダパスが渡される
- 空文字列の場合は、デフォルトの `xls-files` フォルダを使用

## トラブルシューティング

**「フォルダーが見つかりません」エラーが表示される場合**

- パスが正しく入力されているか確認
- フォルダが存在するか確認
- ネットワークドライブの場合、接続状態を確認

**マクロが実行されない場合**

- マクロのセキュリティ設定を確認
- ボタンに正しいマクロが登録されているか確認
- VBA Editor でマクロが存在するか確認
