# PDF2Excel

- 作成日: 2026-03-12 23:14 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-13

Excel(M365) の Power Query を使って、複数のテキストPDFをまとめて Excel に変換するローカルツールです。  
追加インストールなしで、`BAT + PowerShell + Excel` のみで動く構成にしています。

詳しい使い方は [ユーザーマニュアル](C:\Work_Codex\PDF2Excel\docs\user-manual.md) を参照してください。
フォルダ構成は [project-layout.md](C:\Work_Codex\PDF2Excel\docs\project-layout.md) を参照してください。

## はじめに

最初に使うときは、[run_pdf2excel.bat](C:\Work_Codex\PDF2Excel\run_pdf2excel.bat) をダブルクリックしてください。  
画面に出る番号メニューから選ぶだけで変換できます。

- `1`: PDFファイルを複数選んで変換
- `2`: PDFフォルダを選んで変換
- `3`: 使い方マニュアルを開く
- `4`: 出力フォルダを開く
- `5`: 終了

保存先を指定しない場合は、`output` フォルダに `PDF2Excel_yyyyMMdd_HHmmss.xlsx` が作成されます。

## 構成

- `run_pdf2excel.bat`
  - 利用者向けの起動入口です。対話メニュー付きです。
- `docs/`
  - 利用マニュアルとフォルダ構成ガイドを置いています。
- `samples/pdf/`
  - 動作確認用のサンプル PDF です。
- `scripts/run_pdf2excel.ps1`
  - PDF の staging、Excel 起動、Power Query 更新、xlsx 出力を行います。
- `scripts/build_excel_template.ps1`
  - `template/PDF2Excel_Converter.xlsm` を自動生成します。
- `template/vba/PDF2ExcelMacros.bas`
  - Excel テンプレートへ取り込む VBA モジュールです。
- `input/`
  - 処理対象PDFを一時配置する staging フォルダです。
- `output/`
  - 出力された `xlsx` を保存します。
- `logs/`
  - 実行ログを保存します。
- `output/runtime/`
  - 実行中の一時 `xlsm` を置きます。通常は実行ごとに自動クリーンアップされます。
- `tests/`
  - 統合テストとその一時生成物を置きます。
- `reports/`
  - 最新のテスト結果レポートを置きます。

## ディレクトリマップ

```text
PDF2Excel
├─ run_pdf2excel.bat
├─ README.md
├─ docs/
├─ samples/pdf/
├─ scripts/
├─ template/
├─ tests/
├─ reports/
├─ input/
├─ output/
└─ logs/
```

## 使い方

1. [run_pdf2excel.bat](C:\Work_Codex\PDF2Excel\run_pdf2excel.bat) を実行します。
2. メニューで `1` または `2` を選びます。
3. PDF または PDF フォルダを選択します。
4. 保存先を選びます。
5. 処理完了後、`Result` シートと `Errors` シートを確認します。

PowerShell から実行する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -SelectInputFolder -PromptForOutputFile
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFiles "C:\PDF\a.pdf,C:\Other\b.pdf" -OutputFile C:\Work\result.xlsx
```

## 出力仕様

- `Control` シート
  - 実行時の入力フォルダ、出力先、ログファイル、最終実行状態を保持します。
  - あわせて、対象PDF数、取込データ行数、エラー件数を表示します。
- `Result` シート
  - `A列=SourceFile`
  - `B〜AE列=Column1〜Column30`
- `Errors` シート
  - 抽出失敗したPDFのファイル名、失敗理由、候補表の列数・行数を出力します。

## 前提条件

- Windows 上の Excel(M365) デスクトップ版が利用可能であること
- PDF がテキストPDFであること
- 50件程度の PDF が同一または近いレイアウトの表であること
- 1ファイルにつき主対象の表が1種類であること

## 補足

- Power Query の `Pdf.Tables` を使うため、画像PDFやOCR前提のPDFは対象外です。
- 列数が 30 を超えると `Errors` シートへ退避します。
- 列数が 30 未満のときは空列を補完して 30 列へ揃えます。
- マクロはテンプレート作成時に自動で取り込みます。PowerShell 側でも同等の fallback 処理を持たせているため、マクロ実行に失敗しても処理継続できる設計です。
- 同名PDFを別フォルダから同時投入する運用は非対応です。ファイル名が衝突した場合はエラーで止めます。

## テスト

統合テストは [run_integration_tests.ps1](C:\Work_Codex\PDF2Excel\tests\run_integration_tests.ps1) で実行できます。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1
```

結果は次に出力されます。

- レポート: [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md)
- JSON: [integration-test-results.json](C:\Work_Codex\PDF2Excel\tests\results\integration-test-results.json)
