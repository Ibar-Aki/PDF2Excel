# PDF2Excel

- 作成日: 2026-03-12 23:14 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-12

Excel(M365) の Power Query を使って、複数のテキストPDFをまとめて Excel に変換するローカルツールです。  
追加インストールなしで、`BAT + PowerShell + Excel` のみで動く構成にしています。

## 構成

- `run_pdf2excel.bat`
  - 利用者向けの起動入口です。
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

## 使い方

1. `run_pdf2excel.bat` を実行します。
2. PDFファイル選択ダイアログで対象PDFを複数選択します。  
   または PowerShell で `scripts/run_pdf2excel.ps1 -InputFolder <PDFフォルダ>` を実行します。
3. 初回実行時は `template/PDF2Excel_Converter.xlsm` が自動生成されます。
4. 処理完了後、`output/PDF2Excel_yyyyMMdd_HHmmss.xlsx` が出力されます。
5. `InputFolder` に `input/` 自体を指定しても動作します。

## 出力仕様

- `Control` シート
  - 実行時の入力フォルダ、出力先、ログファイル、最終実行状態を保持します。
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
