# PDF2Excel VER2 最小配布パッケージ

- 作成日: 2026-03-16 03:20 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-19

## これは何か

このフォルダは、PDF2Excel `VER2 Secure` を他の人へ渡すための最小構成です。  
生データ転記サンプルを前提に、実行に必要なものだけを残しています。

## VER2 の特徴

- 生データ転記は、PDF から読めた値を、意味解釈や標準化を最小限にして Excel に近い形で出力する方式です。
- `VER2` は secure 既定です。
- 実行中の `runtime / staging` は `%LOCALAPPDATA%\PDF2Excel\runtime\runs` を使います。
- 既定ログ保存先は `%LOCALAPPDATA%\PDF2Excel\logs` です。
- `VER2 Secure` のログは、既定で詳細パスをマスクして記録します。
- 今回選んだ PDF は `input` フォルダへ複製しません。
- 共有パス上からの `VER2 Secure` 実行は拒否します。ローカルへ展開して使ってください。
- 正式運用では、この ZIP をローカルへ展開して `run_pdf2excel.bat` だけを使ってください。
- 出力された `xlsx` だけを部署共有へ移動し、テンプレートやスクリプトは共有フォルダ上で直接更新しないでください。
- 起動導線は `RemoteSigned` 前提です。
- 出力 Excel には `Review` シートが追加されます。

## 同梱ファイル

- `run_pdf2excel.bat`
  - もっとも簡単な起動方法です。
- `scripts/run_pdf2excel.ps1`
  - 変換処理の本体です。
- `scripts/run_pdf2excel_menu.ps1`
  - 利用者向けメニューです。
- `scripts/new_profile_scaffold.ps1`
  - `VER2` のプロファイル雛形を作る保守者向けスクリプトです。
- `scripts/pdf2excel.common.ps1`
  - 共通関数です。
- `template/PDF2Excel_V2_Converter.xlsm`
  - `VER2` 用 Excel テンプレートです。
- `template/vba/PDF2ExcelMacros.bas`
  - 空の `xlsm` へ取り込む基本マクロです。
- `template/vba/PDF2ExcelTemplateBuilder.bas`
  - 空の `xlsm` へ取り込んで実行すると、`VER2` を含むテンプレートのシート構成を自動作成します。
- `config/profiles/v2/construction_transfer_poc.json`
  - 既定の `VER2` プロファイルです。
- `input`
  - 空フォルダです。`VER2 secure` では今回 PDF の保管先として使いません。
- `output`
  - 出力された Excel の保存先です。
- `logs`
  - 互換用に残している空フォルダです。`VER2 Secure` の既定ログ保存先は `%LOCALAPPDATA%\PDF2Excel\logs` です。

## 使う前の条件

- Windows 環境であること
- Excel(M365) デスクトップ版が使えること
- PDF がテキスト選択できること
- 近いレイアウトの帳票をまとめて処理すること

## いちばん簡単な使い方

1. `run_pdf2excel.bat` をダブルクリックします。
2. メニューで `1` または `2` を選びます。必要なら `6` でプロファイル雛形を作れます。
3. PDF または PDF フォルダを選びます。
4. 保存先を選びます。
5. 実行前チェックを確認し、問題なければ続行します。
6. 完了後に、出力された Excel の `Summary`、`Result`、`Review`、`Errors` を確認します。
7. 必要なら、完成した `xlsx` だけを部署共有へ移動します。

## 空の Excel からテンプレートを作る

1. 空の Excel ブックを `xlsm` 形式で保存します。
2. VBA エディターを開き、`template/vba/PDF2ExcelMacros.bas` と `template/vba/PDF2ExcelTemplateBuilder.bas` を標準モジュールとして取り込みます。
3. `VER2` テンプレートを作る場合は `BuildPDF2ExcelV2TemplateInActiveWorkbook` を実行します。
4. `Control / Result / Errors / Summary / Review` が自動作成されたら、`template/PDF2Excel_V2_Converter.xlsm` として保存または置き換えます。

文字化けする環境では、同じ名前の `*.sjis.bas` を代わりに使ってください。

## 出力の見方

### `Summary`

- 全体件数
- 成功 PDF 数
- 失敗 PDF 数
- PDF ごとの取込件数
- エラー分類別件数

### `Result`

- 成功したデータが入ります。
- `VER2` では正規化時刻列も追加されます。

### `Review`

- 確認が必要な行が出ます。
- `ReasonCategory`、`時刻正規化状態`、`時刻確認メモ` を確認してください。

### `Errors`

- 失敗した PDF が出ます。
- `ErrorCode`
- `ErrorCategory`
- `UserMessage`

を確認してください。

## よくある注意点

- 同じ名前の PDF を同時に処理しないでください。
- 画像 PDF やスキャン PDF は対象外です。
- レイアウトが大きく違う PDF を混ぜると `Errors` や `Review` が増えます。
- 既定では `config/profiles/v2/construction_transfer_poc.json` を使います。

## 配布時のおすすめ

- このフォルダごと渡してください。
- 配布先では ZIP をローカルへ展開してから使ってください。
- 共有フォルダ上の ZIP や展開済みフォルダからは直接実行しないでください。
- まずは 2〜3 件で試してから本番件数へ広げてください。

## 困ったとき

- `%LOCALAPPDATA%\PDF2Excel\logs` の最新ログを見る
- `Errors` シートを見る
- `Review` シートを見る
- 実行前チェックの入力元、保存先、件数が正しいか確認する
