# PDF2Excel 最小配布パッケージ

- 作成日: 2026-03-14 00:05 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-15

## これは何か

このフォルダは、PDF2Excel を他の人へ渡すための最小構成です。  
開発用ファイル、テスト、サンプル、詳細レポートは含めず、実行に必要なものだけをまとめています。

## 同梱ファイル

- `run_pdf2excel.bat`
  - もっとも簡単な起動方法です。
- `scripts/run_pdf2excel.ps1`
  - 変換処理の本体です。
- `template/PDF2Excel_Converter.xlsm`
  - 実行に使う Excel テンプレートです。
- `config/profiles/default.json`
  - 既定の帳票プロファイルです。
- `config/profiles/attendance_monthly_jp.json`
  - 日本語の月次勤怠管理表向けプロファイルです。
- `input`
  - 一時的な PDF の配置先です。
- `output`
  - 出力された Excel の保存先です。
- `logs`
  - 実行ログの保存先です。

## 使う前の条件

- Windows 環境であること
- Excel(M365) デスクトップ版が使えること
- PDF がテキスト選択できること
- 同じようなレイアウトの帳票をまとめて処理すること

## いちばん簡単な使い方

1. `run_pdf2excel.bat` をダブルクリックします。
2. メニューで `1` または `2` を選びます。
3. PDF または PDF フォルダを選びます。
4. 保存先を選びます。
5. 実行前チェックを確認し、問題なければ続行します。
6. 完了後に、出力された Excel の `Summary`、`Result`、`Errors` を確認します。

## 出力の見方

### `Summary`

- 全体件数
- 成功 PDF 数
- 失敗 PDF 数
- PDF ごとの取込件数
- エラー分類別件数

### `Result`

- 成功したデータが入ります。
- A列は元の PDF ファイル名です。
- B列以降は抽出された表データです。

### `Errors`

- 失敗した PDF が出ます。
- `ErrorCode`
- `ErrorCategory`
- `UserMessage`
- `TechnicalDetail`

を確認してください。

## よくある注意点

- 同じ名前の PDF を同時に処理しないでください。
- 画像 PDF やスキャン PDF は対象外です。
- レイアウトが大きく違う PDF を混ぜると `Errors` が増えます。
- 既定では `config/profiles/default.json` を使います。
- 日本語の勤怠管理表を扱う場合は `config/profiles/attendance_monthly_jp.json` も試してください。

## 配布時のおすすめ

- このフォルダごと渡してください。
- 配布先で zip を展開したあと、そのまま使えます。
- まずは 2〜3 件で試してから本番件数へ広げてください。

## 困ったとき

- `logs` フォルダの最新ログを見る
- `Errors` シートを見る
- 実行前チェックの入力元、保存先、件数が正しいか確認する
