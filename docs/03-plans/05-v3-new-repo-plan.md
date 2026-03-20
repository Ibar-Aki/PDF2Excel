# PDF2Excel V3 新規レポジトリ作成計画

- 作成日: 2026-03-20 15:01 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## 目的

この文書は、現在の `V2` 実装をベースにしつつ、`V3` を別レポジトリとして新規作成するための計画書です。  
目的は機能追加よりも、保守運用性、責務分離、再配布性、テストしやすさを高めることです。

## 前提

- `V3` は現行 repo の上書き改修ではなく、新規レポジトリとして作成します。
- 技術スタックは変えません。
  - `BAT`
  - `PowerShell 5.1`
  - `Excel(M365)`
  - `Power Query`
  - `VBA`
  - `JSON`
- `V2` から流用できるものは、`V3` 側ディレクトリへコピーして移植します。
- `V2` は別保管とし、`V3` の最終コード・文書・配布物には `V2` の名称、互換分岐、旧命名を残しません。
- `V1` は `V3` へ持ち込みません。
- `Shift_JIS` ミラーの VBA モジュールは、配布要件があるため `V3` でも維持します。

## 結論

`V3` は「機能を大きく増やす版」ではなく、「現行 `V2 Secure` の実運用核を抽出し、`V3` 専用構成へ再編した保守運用版」として作るのが妥当です。  
そのため、現行 repo から必要なファイルをコピーした後、`V3` repo では `V2` 名や互換分岐を残さず、最終的に `V3` 専用コードベースへ整理するべきです。

## V3 の狙い

1. `run_pdf2excel.ps1` への責務集中を解消する
2. 利用者入口と保守者入口を分離する
3. 正本、生成物、実行生成物を明確に分離する
4. テスト対象とモジュール境界を対応させる
5. handoff とテンプレート再生成を標準手順として固定する

## V3 で持ち込むもの

### 初期コピー元として使ってよいもの

- `run_pdf2excel.bat`
- `scripts/run_pdf2excel.ps1`
- `scripts/run_pdf2excel_menu.ps1`
- `scripts/new_profile_scaffold.ps1`
- `scripts/pdf2excel.common.ps1`
- `scripts/build_excel_template.ps1`
- `scripts/build_handoff_package.ps1`
- `scripts/build_sample_pdfs.ps1`
- `template/PDF2Excel_V2_Converter.xlsm`
- `template/vba/*.bas`
- `config/profiles/v2/*`
- `config/template-integrity.json`
- `samples/common/*`
- `samples/v2/*`
- `tests/run_unit_tests.ps1`
- `tests/run_integration_tests.ps1`
- 利用者向け文書のうち現行導線説明に使えるもの

### コピー後に必ず改名・再編するもの

- `run_pdf2excel_v2.bat`
  - `run_pdf2excel_v3.bat` へ改名
- `scripts/run_pdf2excel_v2.ps1`
  - `app/entrypoints/run_pdf2excel_v3.ps1` または `app/entrypoints/run_pdf2excel.ps1` へ改名
- `template/PDF2Excel_V2_Converter.xlsm`
  - `assets/template/PDF2Excel_V3_Converter.xlsm` へ改名
- `config/profiles/v2`
  - `config/profiles/v3` へ改名
- `samples/v2`
  - `assets/samples/v3` へ改名
- `docs/02-guides/01-v2-sample-walkthrough.md`
  - `docs/v3-sample-walkthrough.md` へ改名
- 画面文言、README、handoff README 内の `VER2` / `V2` 表記
  - すべて `V3` 表記へ置換

### 持ち込むが配置を変えるもの

- `scripts/*`
  - `app/entrypoints`, `app/core`, `app/modules`, `tools/build` に分解する
- `template/*`
  - `assets/template` へ移す
- `samples/common`, `samples/v2`
  - `assets/samples/common`, `assets/samples/v2` へ移す
- `tests/*`
  - `tests/unit`, `tests/integration`, `tests/helpers`, `tests/fixtures` に再配置する
- `handoff/sources`, `handoff/generated`
  - 構成は維持するが、新 repo の標準構成として最初から採用する

### 持ち込まないもの

- `legacy/v1/*`
- `V1` 専用 BAT / PowerShell ラッパー
- `V1` 専用プロファイル
- `V1` 専用サンプル
- `V1` handoff
- `V1` 前提の説明文書
- `V2` 互換のためだけに残っている分岐
- `V2` という版名を前提にした命名
- `V2` 固有の説明文書

## V3 推奨リポジトリ名

- `PDF2ExcelV3`
- `pdf2excel-v3`

命名はどちらでもよいですが、PowerShell や ZIP 配布物の見通しを考えると `PDF2ExcelV3` の方が扱いやすいです。

## V3 推奨フォルダ構成

```text
PDF2ExcelV3
├─ run_pdf2excel.bat
├─ run_pdf2excel_v3.bat
├─ README.md
├─ docs/
│  ├─ index.md
│  ├─ user-manual.md
│  ├─ maintenance-guide.md
│  ├─ project-layout.md
│  ├─ technical-description.md
│  ├─ v3-sample-walkthrough.md
│  └─ import-history-from-v2.md
├─ app/
│  ├─ entrypoints/
│  │  ├─ run_pdf2excel.ps1
│  │  ├─ run_menu.ps1
│  │  ├─ run_environment_check.ps1
│  │  └─ new_profile.ps1
│  ├─ core/
│  │  ├─ orchestrator.psm1
│  │  ├─ pipeline.psm1
│  │  └─ execution-context.psm1
│  ├─ modules/
│  │  ├─ logging.psm1
│  │  ├─ filesystem.psm1
│  │  ├─ locking.psm1
│  │  ├─ runtime.psm1
│  │  ├─ environment-check.psm1
│  │  ├─ profiles.psm1
│  │  ├─ profile-wizard.psm1
│  │  ├─ pdf-staging.psm1
│  │  ├─ excel-session.psm1
│  │  ├─ powerquery.psm1
│  │  ├─ result-shaping.psm1
│  │  ├─ review-output.psm1
│  │  └─ reports.psm1
│  └─ contracts/
│     ├─ run-report.schema.json
│     ├─ environment-check.schema.json
│     └─ profile-v3.schema.json
├─ config/
│  ├─ app/
│  │  ├─ defaults.json
│  │  └─ paths.json
│  ├─ profiles/
│  │  └─ v3/
│  └─ template-integrity.json
├─ assets/
│  ├─ template/
│  │  ├─ PDF2Excel_V3_Converter.xlsm
│  │  └─ vba/
│  │     ├─ PDF2ExcelMacros.bas
│  │     ├─ PDF2ExcelMacros.sjis.bas
│  │     ├─ PDF2ExcelTemplateBuilder.bas
│  │     └─ PDF2ExcelTemplateBuilder.sjis.bas
│  └─ samples/
│     ├─ common/
│     └─ v3/
├─ tools/
│  ├─ build/
│  │  ├─ build_excel_template.ps1
│  │  ├─ build_handoff_package.ps1
│  │  └─ build_sample_pdfs.ps1
│  ├─ dev/
│  │  ├─ validate_repo_layout.ps1
│  │  ├─ validate_profiles.ps1
│  │  └─ sync_sjis_mirror.ps1
│  └─ migration/
│     └─ import_from_v2.ps1
├─ handoff/
│  ├─ sources/
│  └─ generated/
├─ tests/
│  ├─ unit/
│  ├─ integration/
│  │  ├─ smoke/
│  │  └─ full/
│  ├─ helpers/
│  ├─ fixtures/
│  ├─ results/
│  └─ work/
├─ reports/
│  ├─ README.md
│  ├─ tracked/
│  └─ runtime/
├─ input/
├─ output/
└─ logs/
```

## V3 の設計方針

### 1. 入口を薄くする

- `BAT` は利用者向け入口だけに絞る
- `entrypoints` は引数解決とモジュール呼び出ししかしない
- 実処理は `app/core` と `app/modules` に寄せる

### 2. PowerShell をモジュール分割する

- 現行の巨大スクリプトを `.psm1` 単位へ分割する
- 1 モジュール 1 責務を基本にする
- 単体テストが関数単位でかけられる形にする

### 3. 正本と生成物を分離する

- 編集するものは `assets` と `config`
- 再生成スクリプトは `tools/build`
- 実行時生成物は `reports/runtime`
- 配布物は `handoff/generated`

### 4. V3 は最終的に V2 痕跡を残さない

- `V1` 互換維持コードを持ち込まない
- メニュー、文書、handoff、サンプルは `V3` 専用にする
- `V2` の「共通化のための分岐」は残さない
- `V2` 命名の BAT、PowerShell、テンプレート、プロファイル、サンプル名は最終的にすべて `V3` へ置換する
- コメント、ログ、レポート文言に残る `V2` 名も整理対象に含める

## V2 から V3 へのコピー方針

### 方針

- 元 repo のファイルは移動しない
- `V3` repo を新規作成し、必要ファイルだけコピーする
- 初期段階では動作確認を優先してコピーする
- その後、`V3` repo 側では `V2` 命名と互換分岐を完全に除去する

### 推奨する初期コピー単位

1. 利用者入口
   - `run_pdf2excel.bat`
2. 実行本体
   - `scripts/run_pdf2excel.ps1`
   - `scripts/run_pdf2excel_menu.ps1`
   - `scripts/new_profile_scaffold.ps1`
   - `scripts/pdf2excel.common.ps1`
3. テンプレート
   - `template/`
   - `config/template-integrity.json`
4. プロファイル
   - `config/profiles/v2/`
5. サンプル
   - `samples/common/`
   - `samples/v2/`
6. テスト
   - `tests/run_unit_tests.ps1`
   - `tests/run_integration_tests.ps1`
7. 再生成系
   - `scripts/build_excel_template.ps1`
   - `scripts/build_handoff_package.ps1`
   - `scripts/build_sample_pdfs.ps1`
8. 文書
   - `README.md`
   - `docs/index.md`
   - `docs/01-current/01-user-manual.md`
   - `docs/01-current/05-maintenance-guide.md`
   - `docs/01-current/03-project-layout.md`
   - `docs/01-current/04-technical-description.md`
   - `docs/02-guides/01-v2-sample-walkthrough.md`

### コピー直後に行う命名整理

1. `run_pdf2excel_v2.bat` を `run_pdf2excel_v3.bat` へ改名
2. `scripts/run_pdf2excel_v2.ps1` を `app/entrypoints/run_pdf2excel_v3.ps1` へ改名
3. `template/PDF2Excel_V2_Converter.xlsm` を `assets/template/PDF2Excel_V3_Converter.xlsm` へ改名
4. `config/profiles/v2` を `config/profiles/v3` へ改名
5. `samples/v2` を `assets/samples/v3` へ改名
6. `docs/02-guides/01-v2-sample-walkthrough.md` を `docs/v3-sample-walkthrough.md` へ改名
7. コード、コメント、ログ、README、handoff README の `V2` 表記を `V3` へ置換

## V3 作成フェーズ

### Phase 1. 新規 repo 作成と最小起動

- `V3` repo を新規作成
- 現行 `V2` の起動導線とテンプレートをコピー
- `V1` を持ち込まずに `smoke` 相当が通る状態を作る
- ただし Phase 1 の終わりまでに `V2` 命名の入口ファイルは `V3` 名へ改名する

### Phase 2. フォルダ再配置

- `scripts` を `app/entrypoints`, `app/core`, `app/modules`, `tools/build` へ分ける
- `template` を `assets/template` へ移す
- `samples` を `assets/samples` へ移す
- テストを `tests/unit`, `tests/integration` へ再編する
- `config/profiles/v2` を `config/profiles/v3` へ切り替える

### Phase 3. モジュール分割

- `logging`
- `locking`
- `environment-check`
- `profiles`
- `reports`
- `excel-session`
- `powerquery`
- `result-shaping`

この順で `run_pdf2excel.ps1` から外出しします。

### Phase 4. 文書再整備

- `V3` 用 README を新規作成
- `V2` 由来の説明文言を整理
- `V3` の正式運用手順、handoff 手順、障害対応を再記述
- `V2` という版名を残す文書を repo 内から除去する

### Phase 5. 配布と回帰確立

- `build_handoff_package.ps1` を `V3` 用へ調整
- `smoke` と `full` の 2 スイートで回帰を固定
- `reports/tracked` と `reports/runtime` の運用を固める

## 推定作業時間: 約32時間

## V3 作成時の受け入れ条件

1. `V3` repo 単体で `run_pdf2excel.bat` から起動できる
2. `V1` に依存しない
3. `smoke` テストが通る
4. handoff を `V3` repo 単体で再生成できる
5. テンプレートを `V3` repo 単体で再生成できる
6. `Shift_JIS` ミラー VBA を配布物へ含められる
7. 利用者向け文書と保守向け文書が分離されている
8. repo 内のコード、設定、文書、配布物に `V2` 互換や `V2` 命名が残っていない

## リスクと注意点

- 現行 `V2` の分岐には、`V1` 互換や `V2` 命名の名残が一部残っている可能性があります。
- 新 repo 初期段階では、まず「動く状態のコピー」を優先し、その後で `V3` 命名への全面置換を行うべきです。
- 最初から全面リファクタリングすると、Excel COM と Power Query の不具合切り分けが難しくなります。
- テンプレート再生成と handoff 再生成は、初期フェーズから必ず回帰対象に含めるべきです。

## 推奨する最初の一歩

1. `PDF2ExcelV3` repo を作る
2. `V2` 正式導線だけをコピーする
3. `V2` 命名を `V3` 命名へ置換する
4. `smoke` テストが通るまで修正する
5. その後に `app/modules` へ分割を始める

この順なら、機能を壊さずに `V3` へ移行しやすく、最終的に `V2` の痕跡も残りません。
