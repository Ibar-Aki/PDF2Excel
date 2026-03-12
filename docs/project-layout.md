# PDF2Excel フォルダ構成ガイド

- 作成日: 2026-03-13 00:49 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-13

## 目的

このファイルは、`C:\Work_Codex\PDF2Excel` の中身を誰でも迷わず理解できるようにするための案内です。

## ひと目でわかる構成

```text
PDF2Excel
├─ run_pdf2excel.bat        利用者向けの入口
├─ README.md                最初に読む概要
├─ docs/                    マニュアルと構成説明
├─ samples/pdf/             動作確認用のサンプル PDF
├─ scripts/                 PowerShell 本体
├─ template/                Excel テンプレートと VBA
├─ tests/                   統合テスト
├─ reports/                 テスト結果レポート
├─ input/                   実行時に一時配置する PDF
├─ output/                  実行結果の xlsx
└─ logs/                    実行ログ
```

## 利用者が主に触る場所

- [index.md](C:\Work_Codex\PDF2Excel\docs\index.md)
  - 文書の入口です。
- [run_pdf2excel.bat](C:\Work_Codex\PDF2Excel\run_pdf2excel.bat)
  - 変換を始める入口です。
- [README.md](C:\Work_Codex\PDF2Excel\README.md)
  - 全体概要を短く確認できます。
- [user-manual.md](C:\Work_Codex\PDF2Excel\docs\user-manual.md)
  - 詳しい使い方です。
- `samples/pdf`
  - サンプル PDF です。
- `output`
  - 変換後の Excel が出ます。
- `logs`
  - エラー調査時に見ます。

## 保守時に触る場所

- [run_pdf2excel.ps1](C:\Work_Codex\PDF2Excel\scripts\run_pdf2excel.ps1)
  - 変換本体です。
- [build_excel_template.ps1](C:\Work_Codex\PDF2Excel\scripts\build_excel_template.ps1)
  - `xlsm` テンプレートを再生成します。
- [PDF2ExcelMacros.bas](C:\Work_Codex\PDF2Excel\template\vba\PDF2ExcelMacros.bas)
  - Excel 側のマクロです。
- [run_integration_tests.ps1](C:\Work_Codex\PDF2Excel\tests\run_integration_tests.ps1)
  - 統合テストです。
- [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md)
  - 最新のテスト結果です。

## 運用ルール

- `input` は一時置き場です。常設保管には使いません。
- `output` は成果物置き場です。必要なものだけ残してください。
- `logs` は実行ごとに増えるため、不要になったら消して構いません。
- `tests/results` と `tests/work` はテストの生成物です。通常は空で問題ありません。

## おすすめの見方

1. まず [README.md](C:\Work_Codex\PDF2Excel\README.md) を読む
2. 次に [user-manual.md](C:\Work_Codex\PDF2Excel\docs\user-manual.md) を読む
3. 実行は [run_pdf2excel.bat](C:\Work_Codex\PDF2Excel\run_pdf2excel.bat) から始める
4. 問題が出たら `logs` と [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md) を確認する
