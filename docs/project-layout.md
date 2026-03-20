# PDF2Excel フォルダ構成ガイド

- 作成日: 2026-03-13 00:49 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## 目的

このファイルは、PDF2Excel プロジェクトルートの中身を誰でも迷わず理解できるようにするための案内です。

## ひと目でわかる構成

```text
PDF2Excel
├─ run_pdf2excel.bat        正式運用入口 (VER2 Secure)
├─ run_pdf2excel_v2.bat     VER2 を明示起動する入口
├─ README.md                最初に読む概要
├─ docs/                    マニュアルと構成説明
├─ config/                  帳票プロファイルと整合性マニフェスト
├─ handoff/                 配布 README ソースと生成済み配布物
├─ legacy/v1/               VER1 の旧導線・旧サンプル・旧プロファイル
├─ samples/common/          共通サンプルの正本
├─ samples/v2/              VER2 サンプル
├─ scripts/                 PowerShell 本体と共通関数
├─ template/                Excel テンプレートと VBA
├─ tests/                   統合テストとユニットテスト
├─ reports/                 テスト結果と運用レポート
├─ input/                   保管用 PDF
├─ output/                  実行結果の xlsx
└─ logs/                    開発・検証導線の実行ログ
```

## 利用者が主に触る場所

- [index.md](index.md)
  - 文書の入口です。
- [run_pdf2excel.bat](../run_pdf2excel.bat)
  - 正式運用で使う `VER2 Secure` の入口です。
- [run_pdf2excel_v2.bat](../run_pdf2excel_v2.bat)
  - `VER2 Secure` を明示して起動したいときの入口です。
- [README.md](../README.md)
  - 全体概要を短く確認できます。
- [user-manual.md](user-manual.md)
  - 詳しい使い方です。
- `config/profiles/v2`
  - `VER2` 用プロファイルです。
- `samples/v2`
  - `VER2` の生データ転記サンプルです。
- `samples/common`
  - 共通サンプルの正本です。再生成スクリプトはここを更新します。
- `legacy/v1`
  - `VER1` の BAT、PowerShell ラッパー、旧プロファイル、旧サンプル、旧 handoff 定義を置きます。
- `handoff/sources`
  - 配布 README のソース文書です。
- `handoff/generated`
  - 生成済み配布フォルダと ZIP を置きます。
- `output`
  - 変換後の Excel が出ます。
- `%LOCALAPPDATA%\PDF2Excel\logs`
  - `VER2 Secure` の既定ログ保存先です。
- `reports`
  - `run-history.csv` や `environment-check.md` などのレポートが出ます。

## 保守時に触る場所

- [run_pdf2excel.ps1](../scripts/run_pdf2excel.ps1)
  - 変換本体です。
- [run_pdf2excel_v2.ps1](../scripts/run_pdf2excel_v2.ps1)
  - `VER2` ラッパーです。
- [run_pdf2excel_menu.ps1](../scripts/run_pdf2excel_menu.ps1)
  - 日本語の共通対話メニューです。`[9] プロファイル選択` と `[0] 環境チェック` を含みます。
- [run_pdf2excel_menu_v2.ps1](../scripts/run_pdf2excel_menu_v2.ps1)
  - `VER2` 用ラッパーです。内部では共通メニューを呼び出します。
- [new_profile_scaffold.ps1](../scripts/new_profile_scaffold.ps1)
  - プロファイル雛形生成スクリプトです。`VER2` では設定ウィザードを使えます。
- [pdf2excel.common.ps1](../scripts/pdf2excel.common.ps1)
  - 共通関数です。
- [build_excel_template.ps1](../scripts/build_excel_template.ps1)
  - `xlsm` テンプレートと `config/template-integrity.json` を再生成します。
- [build_sample_pdfs.ps1](../scripts/build_sample_pdfs.ps1)
  - 日本語勤怠管理表のサンプル Excel / PDF を再生成します。
- [PDF2ExcelMacros.bas](../template/vba/PDF2ExcelMacros.bas)
  - Excel 側のマクロです。
- [PDF2ExcelMacros.sjis.bas](../template/vba/PDF2ExcelMacros.sjis.bas)
  - Shift_JIS 互換用のミラーです。
- [PDF2ExcelTemplateBuilder.bas](../template/vba/PDF2ExcelTemplateBuilder.bas)
  - 空ブックをテンプレート相当に初期化する VBA です。
- [PDF2ExcelTemplateBuilder.sjis.bas](../template/vba/PDF2ExcelTemplateBuilder.sjis.bas)
  - その Shift_JIS 互換用ミラーです。
- [run_integration_tests.ps1](../tests/run_integration_tests.ps1)
  - 統合テストです。`smoke / full / legacy` の 3 スイートで運用します。
- [run_unit_tests.ps1](../tests/run_unit_tests.ps1)
  - ユニットテストです。
- [test-report.md](../reports/test-report.md)
  - 最新のテスト結果です。
- [unit-test-report.md](../reports/unit-test-report.md)
  - 最新のユニットテスト結果です。
- [README.md](../reports/README.md)
  - `reports/` 配下の追跡対象と生成物を説明します。

## 運用ルール

- `input` は保管置き場です。`VER2 Secure` の正式導線では今回 PDF を `input` へ同期しません。
- `output` は成果物置き場です。必要なものだけ残してください。
- `logs` は主に開発・検証導線で使います。`VER2 Secure` の既定ログは `%LOCALAPPDATA%\PDF2Excel\logs` に出ます。
- `config/profiles` は帳票ごとの設定置き場です。新しい帳票を増やすときはここへ JSON を追加します。
- `legacy/v1` は通常運用で触らない領域です。`VER1` 保守が必要なときだけ参照してください。
- `samples/common` は共通サンプルの正本です。`samples/v2` は利用者導線、`legacy/v1/samples` は互換確認用です。
- `handoff/sources` と `handoff/generated` は混同しないでください。編集対象は source、配布対象は generated です。
- `config/template-integrity.json` はテンプレートと VBA モジュールの整合性確認に使います。
- `reports/run-history.csv` と `reports/environment-check.md` は運用レポートです。追跡対象にしない前提で扱います。
- `tests/results` と `tests/work` はテストの生成物です。通常は空で問題ありません。

## おすすめの見方

1. まず [README.md](../README.md) を読む
2. 次に [user-manual.md](user-manual.md) を読む
3. 実行は [run_pdf2excel.bat](../run_pdf2excel.bat) から始める
4. 問題が出たら [troubleshooting.md](troubleshooting.md)、[README.md](../reports/README.md)、[test-report.md](../reports/test-report.md) を確認する
