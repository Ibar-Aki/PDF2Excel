# PDF2Excel 保守ガイド

- 作成日: 2026-03-13 00:51 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-14

## 目的

この資料は、PDF2Excel を改修、検証、再配布するときの保守作業をまとめたものです。

## 対象ファイル

- [run_pdf2excel.ps1](../scripts/run_pdf2excel.ps1)
  - 変換処理の中核
- [pdf2excel.common.ps1](../scripts/pdf2excel.common.ps1)
  - 共通関数
- [build_excel_template.ps1](../scripts/build_excel_template.ps1)
  - `xlsm` テンプレートの再生成
- [PDF2ExcelMacros.bas](../template/vba/PDF2ExcelMacros.bas)
  - Excel マクロ
- [PDF2ExcelMacros.sjis.bas](../template/vba/PDF2ExcelMacros.sjis.bas)
  - Excel マクロの Shift_JIS ミラー
- [run_integration_tests.ps1](../tests/run_integration_tests.ps1)
  - 統合テスト
- [run_unit_tests.ps1](../tests/run_unit_tests.ps1)
  - ユニットテスト

## 保守の基本方針

- 利用者の入口は `run_pdf2excel.bat` を維持する
- 実行時生成物は `input`、`output`、`logs`、`tests/results` に閉じ込める
- ドキュメント更新をコード変更と同じタイミングで行う
- 帳票認識精度の変更は、実PDFを使った目視確認まで行う

## よくある改修と見るべき場所

### 1. 画面文言や導線を変えたい

見るべき場所:

- [run_pdf2excel.bat](../run_pdf2excel.bat)
- [run_pdf2excel.ps1](../scripts/run_pdf2excel.ps1)
- [README.md](../README.md)
- [user-manual.md](user-manual.md)

注意:

- 対話実行と非対話実行の両方を壊さないこと
- BAT の文字コード制約があるため、文言変更は慎重に行うこと

### 2. Excel テンプレートを変えたい

見るべき場所:

- [build_excel_template.ps1](../scripts/build_excel_template.ps1)
- [PDF2ExcelMacros.bas](../template/vba/PDF2ExcelMacros.bas)
- [PDF2ExcelMacros.sjis.bas](../template/vba/PDF2ExcelMacros.sjis.bas)

手順:

1. `build_excel_template.ps1` または VBA を修正する
2. テンプレートを再生成する
3. `template/PDF2Excel_Converter.xlsm` と `PDF2ExcelMacros.sjis.bas` の更新を確認する
4. ユニットテストと統合テストを実行する

テンプレート再生成:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_excel_template.ps1
```

### 3. PDF 抽出ルールを変えたい

見るべき場所:

- [run_pdf2excel.ps1](../scripts/run_pdf2excel.ps1)
  - `Get-StagingQueryFormula`
  - `Get-ResultQueryFormula`
  - `Get-ErrorsQueryFormula`

注意:

- 30列固定仕様を崩すと、`Result` シートの形もテストも崩れる
- `Errors` に逃がす条件を弱めすぎると、列ズレのまま成功扱いになる

## テスト手順

統合テスト:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1
```

ユニットテスト:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_unit_tests.ps1
```

確認すべき観点:

- 正常な 2 PDF 変換
- 1 PDF 変換
- `-InputFiles` 指定
- 日本語ファイル名
- 50件一括変換
- BAT 実行
- `input` 自己参照
- 同名PDF拒否
- 壊れた PDF 混在
- 深い出力先
- runtime 清掃
- Excel プロセス残留なし

## 変更後チェックリスト

コード変更後は、原則として次を確認します。

1. テンプレート再生成が通る
2. ユニットテストと統合テストが成功する
3. `reports/test-report.md` と `reports/unit-test-report.md` が更新される
4. README と関連ドキュメントが更新されている
5. `input`、`logs`、`tests/results` に不要な生成物が残っていない

## リリース前の確認

- `git status` が意図した差分だけになっている
- README の導線が最新
- `docs/index.md` のリンクが切れていない
- 利用者向け文書と保守向け文書の内容が矛盾していない
- 直近のテストレポートが成功になっている

## 実PDFを使う確認

テストPDFだけでは不足するため、帳票の意味が変わる改修では実PDFでも確認します。

最低限見ること:

- 空欄が多い行で列ズレしないか
- `Errors` に過剰退避していないか
- A列にファイル名が正しく入るか
- 1件だけの PDF でも変換できるか

## ドキュメント更新ルール

- 新規 `*.md` は作成日と作成者を入れる
- 既存 `*.md` は更新日を更新する
- UI/UX に関わる変更は [README.md](../README.md) と [user-manual.md](user-manual.md) に反映する

## 迷ったときの優先順位

1. 実害のある不具合を止める
2. データ破損や誤変換を防ぐ
3. テストで再発防止する
4. その後に使いやすさや文書を整える
