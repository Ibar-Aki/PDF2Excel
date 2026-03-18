# PDF2Excel

- 作成日: 2026-03-12 23:14 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-18

Excel(M365) の Power Query を使って、複数のテキストPDFをまとめて Excel に変換するローカルツールです。  
追加インストールなしで、`BAT + PowerShell + Excel` のみで動く構成にしています。

関連ドキュメント:

- 文書一覧: [index.md](docs/index.md)
- 詳しい使い方: [user-manual.md](docs/user-manual.md)
- 構成説明: [project-layout.md](docs/project-layout.md)
- 障害対応: [troubleshooting.md](docs/troubleshooting.md)
- 保守手順: [maintenance-guide.md](docs/maintenance-guide.md)
- 改善提案: [improvement-proposals.md](docs/improvement-proposals.md)
- 変更履歴: [CHANGELOG.md](CHANGELOG.md)
- Copilot 実装委任: [copilot-implementation-report.md](docs/copilot-implementation-report.md)
- 要件定義書: [requirements-specification.md](docs/requirements-specification.md)
- 技術説明書: [technical-description.md](docs/technical-description.md)
- 生データ転記案: [v2-data-transfer-proposal.md](docs/v2-data-transfer-proposal.md)

詳しい使い方は [ユーザーマニュアル](docs/user-manual.md) を参照してください。
フォルダ構成は [project-layout.md](docs/project-layout.md) を参照してください。

## はじめに

最初に使うときは、用途に応じて次のどちらかを実行してください。  

- [run_pdf2excel_v1.bat](run_pdf2excel_v1.bat)
  - 既存の標準変換です。単一表を素直に Excel 化したいときに使います。
- [run_pdf2excel_v2.bat](run_pdf2excel_v2.bat)
  - 生データ転記です。PDF から読めた値を、意味解釈や標準化を最小限にして Excel に近い形で出力します。
  - 複数ページで同じ列が続く帳票や、確認作業を前提にした転記に使います。
  - `sameHeader` の厳格判定で多ページ結合を行い、曖昧な候補は `Errors` / `Review` に分離します。
  - `VER2` は secure 既定です。runtime / staging は `%LOCALAPPDATA%\PDF2Excel\runtime` 配下を使い、`input` フォルダへ今回 PDF を同期しません。
  - `VER2 Secure` の既定ログ保存先は `%LOCALAPPDATA%\PDF2Excel\logs` です。
  - `VER2 Secure` は共有パス上からの実行を拒否します。ローカルへ展開して使ってください。
  - ダブルクリック起動では、必ず最初にメニューを表示します。
- [run_pdf2excel.bat](run_pdf2excel.bat)
  - 互換入口です。`VER1` を起動します。

画面に出る番号メニューから選ぶだけで変換できます。

- `1`: PDFファイルを複数選んで変換
- `2`: PDFフォルダを選んで変換
- `3`: 使い方マニュアルを開く
- `4`: 出力フォルダを開く
- `5`: プロファイルフォルダを開く
- `6`: プロファイル雛形を作成
- `7`: ログフォルダを開く
- `8`: 終了

保存先を指定しない場合は、`output` フォルダに `PDF2Excel_yyyyMMdd_HHmmss.xlsx` が作成されます。
変換前には、対象件数、保存先、使用プロファイル、想定列数を確認する「実行前チェック」が表示されます。`VER2 Secure` のログは `%LOCALAPPDATA%\PDF2Excel\logs` に保存されます。

## 構成

- `run_pdf2excel_v1.bat` / `run_pdf2excel_v2.bat`
  - 利用者向けの起動入口です。ASCII の起動ラッパーで、対話メニュー本体は PowerShell 側で表示します。
- `run_pdf2excel.bat`
  - 互換入口です。`VER1` を起動します。
- `scripts/run_pdf2excel_menu.ps1`
  - 日本語の共通対話メニューです。版別ラッパーから呼び出されます。
- `scripts/new_profile_scaffold.ps1`
  - v1/v2 のプロファイル雛形を生成する保守者向けスクリプトです。
- `scripts/run_pdf2excel_v1.ps1` / `scripts/run_pdf2excel_v2.ps1`
  - 共通コア `run_pdf2excel.ps1` を版別設定付きで起動するラッパーです。
- `docs/`
  - 利用マニュアルとフォルダ構成ガイドを置いています。
- `samples/v1/`
  - `VER1` 用のサンプル PDF と元 Excel です。
- `samples/v2/`
  - `VER2` 用の生データ転記サンプルです。`2ページ同一列`、`6ページ同一列`、`ヘッダー不一致負例`、`時刻確認負例` を含みます。
- `scripts/run_pdf2excel.ps1`
  - PDF の staging、Excel 起動、Power Query 更新、xlsx 出力を行います。
- `scripts/build_excel_template.ps1`
  - `template/PDF2Excel_V1_Converter.xlsm` と `template/PDF2Excel_V2_Converter.xlsm` を自動生成します。
- `template/vba/PDF2ExcelMacros.bas`
  - Excel テンプレートへ取り込む VBA モジュールです。
- `template/vba/PDF2ExcelTemplateBuilder.bas`
  - 空の Excel ブックに取り込んで実行すると、テンプレート相当のシート構成と基本マクロを生成するブートストラップ用 VBA モジュールです。
- `config/profiles/v1/`
  - `VER1` 用プロファイルです。既定の `default.json` と、日本語の `勤怠管理表`、`売上日報`、`在庫一覧`、`問い合わせ管理表` を置きます。
- `config/profiles/v2/`
  - `VER2` 用プロファイルです。生データ転記用の `construction_transfer_poc.json` を置きます。
- `input/`
  - 処理対象PDFの保管先です。`VER1` と標準導線では、実行時の抽出を `output/runtime/runs/.../staging` で分離して行います。
- `output/`
  - 出力された `xlsx` を保存します。
- `logs/`
  - 主に `VER1` と標準導線の実行ログを保存します。30日超または200件超の古いログは自動整理されます。
  - `VER2 Secure` の既定ログ保存先は `%LOCALAPPDATA%\PDF2Excel\logs` です。
- `output/runtime/`
  - 主に `VER1` と標準導線の実行中一時領域です。`runs/` 配下に実行単位の staging / runtime を作成し、通常は実行ごとに自動クリーンアップされます。
  - `VER2` secure 導線では `%LOCALAPPDATA%\PDF2Excel\runtime\runs` を使います。
- `tests/`
  - 統合テストとその一時生成物を置きます。
- `reports/`
  - 最新のテスト結果レポートを置きます。

## ディレクトリマップ

```text
PDF2Excel
├─ run_pdf2excel.bat
├─ run_pdf2excel_v1.bat
├─ run_pdf2excel_v2.bat
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

1. [run_pdf2excel_v1.bat](run_pdf2excel_v1.bat) または [run_pdf2excel_v2.bat](run_pdf2excel_v2.bat) を実行します。
2. メニューで `1` または `2` を選びます。
3. PDF または PDF フォルダを選択します。
4. 保存先を選びます。
5. 実行前チェックの内容を確認して続行します。
   - `Y` を押した後は `PDF取り込みに時間がかかります。しばらくお待ちください....` が表示されます。
6. 処理完了後、`Summary`、`Result`、`Errors` シートを確認します。
   - 完了画面には出力先とログ先も表示されます。

PowerShell から実行する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -SelectInputFolder -PromptForOutputFile
```

```powershell
powershell -NoProfile -File .\scripts\run_pdf2excel_v2.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result_v2.xlsx
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFiles "C:\PDF\a.pdf,C:\Other\b.pdf" -OutputFile C:\Work\result.xlsx
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -ProfileName default -OutputFile C:\Work\result.xlsx
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -ProfilePath C:\Work\profile10.json -OutputFile C:\Work\result.xlsx -NoConfirm
```

## テンプレートを空ブックから再作成する

空の Excel ブックへ VBA モジュール 2 つを取り込めば、テンプレートをその場で組み立てられます。

1. Excel で空のブックを開き、`xlsm` 形式で保存します。
2. VBA エディターを開き、標準モジュールとして [PDF2ExcelMacros.bas](template/vba/PDF2ExcelMacros.bas) と [PDF2ExcelTemplateBuilder.bas](template/vba/PDF2ExcelTemplateBuilder.bas) を取り込みます。
3. 次のいずれかのマクロを実行します。
   - `VER1`: `BuildPDF2ExcelV1TemplateInActiveWorkbook`
   - `VER2`: `BuildPDF2ExcelV2TemplateInActiveWorkbook`
4. `VER2` では `Control / Result / Errors / Summary / Review` の 5 シートが自動作成されます。
5. `run_pdf2excel` から使う場合は、保存したブックを `template/PDF2Excel_V1_Converter.xlsm` または `template/PDF2Excel_V2_Converter.xlsm` として配置し直します。

既定の互換入口として `BuildPDF2ExcelTemplateInActiveWorkbook` も残しており、これは `VER1` を組み立てます。

## 出力仕様

- `Control` シート
  - 実行時の入力フォルダ、出力先、ログファイル、最終実行状態を保持します。
  - あわせて、対象PDF数、取込データ行数、エラー件数、成功PDF数、失敗PDF数、処理時間、使用プロファイルを表示します。
- `Summary` シート
  - PDFごとの取込件数と成功/失敗の内訳を表示します。
  - エラー分類別件数も表示します。
- `Result` シート
  - `A列=SourceFile`
  - `B列以降=プロファイルに応じた表データ列`
  - `VER2` では生データ列を保持したまま、末尾に `正規化入場*` / `正規化退場*` / `*_分` / `時刻正規化状態` / `時刻確認メモ` を追加します。
- `Errors` シート
  - 抽出失敗したPDFのファイル名、エラーコード、エラー分類、利用者向けメッセージ、技術詳細、候補表の列数・行数を出力します。
- `Review` シート
  - `VER2` のみです。確認が必要な行を、元ファイル名、ページ、氏名 raw、現場 raw、入場 raw、退場 raw、確認要理由で一覧化します。
  - `ReasonCategory` を追加し、`HEADER_MISMATCH`、`TIME_MISSING`、`TIME_MULTI`、`TIME_INVALID` を機械的に判別できるようにしています。
  - あわせて `正規化入場` / `正規化退場` / `*_分` / `時刻正規化状態` / `時刻確認メモ` を出し、`08：00`、`9時 15分`、Excel 時刻比率文字列の確認と後続分析をしやすくします。

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
- 空ブックからテンプレートを作りたい場合は `template/vba/PDF2ExcelTemplateBuilder.bas` または `template/vba/PDF2ExcelTemplateBuilder.sjis.bas` を VBA エディタへインポートし、`BuildPDF2ExcelTemplateInActiveWorkbook` を実行してください。
- `PDF2ExcelTemplateBuilder.*.bas` は空ブック専用です。既存の `PDF2Excel_Converter.xlsm` へ追加インポートすると同名マクロが重複します。
- 同名PDFを別フォルダから同時投入する運用は非対応です。ファイル名が衝突した場合はエラーで止めます。
- `-KeepInput` は `input` フォルダの保管内容を残すためのオプションです。今回の変換対象は毎回専用 staging に切り出して処理するため、過去PDFが混ざることはありません。
- `VER2` secure では `-KeepInput` は無効です。今回 PDF を `input` へ複製せず、ローカル runtime の staging だけで処理します。
- `VER2` の配布用 ZIP と利用者向け BAT / PS 導線は `ExecutionPolicy Bypass` を使わず、`RemoteSigned` を前提に起動します。
- ツールは同時に 1 実行だけ許可します。別実行が動作中の場合は `RUN_LOCKED` として即時停止します。
- BAT メニューと利用者向けの説明文は日本語化しています。
- 日本語の月次勤怠管理表を試す場合は `config/profiles/v1/attendance_monthly_jp.json` を利用してください。
- 日本語の売上日報を試す場合は `config/profiles/v1/sales_daily_jp.json` を利用してください。
- 日本語の在庫一覧を試す場合は `config/profiles/v1/inventory_list_jp.json` を利用してください。
- 日本語の問い合わせ管理表を試す場合は `config/profiles/v1/inquiry_weekly_jp.json` を利用してください。
- 生データ転記サンプルを試す場合は `config/profiles/v2/construction_transfer_poc.json` を利用してください。
- 日本語サンプル PDF を再生成したい場合は `scripts/build_sample_pdfs.ps1` を実行してください。

## テスト

統合テストは [run_integration_tests.ps1](tests/run_integration_tests.ps1) で実行できます。

ユニットテストは [run_unit_tests.ps1](tests/run_unit_tests.ps1) で実行できます。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1
```

結果は次に出力されます。

- レポート: [test-report.md](reports/test-report.md)
- JSON: [integration-test-results.json](tests/results/integration-test-results.json)
- ユニットテストレポート: [unit-test-report.md](reports/unit-test-report.md)

統合テストは `CaseName` を付けなければ全件実行です。  
全件実行では、既知の Excel COM 一時失敗だけを 1 回だけ再試行し、`RetryCount` / `RetriedBy` をレポートへ残します。
