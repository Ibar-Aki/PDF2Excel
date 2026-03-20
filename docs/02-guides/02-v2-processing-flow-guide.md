# V2 処理フロー 完全ガイド

- 作成日: 2026-03-15 21:52 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20
- 目的: PDF 3 件が Excel に転記される仕組みを学習するためのドキュメント

## 全体像

PDF2Excel V2 は **4 層アーキテクチャ** で構成されています。

```
┌───────────────────────────────────────────────┐
│  ① BAT 層      : run_pdf2excel_v2.bat        │ ユーザーがダブルクリック
├───────────────────────────────────────────────┤
│  ② PS 制御層   : run_pdf2excel.ps1           │ 全体の制御・入出力管理
├───────────────────────────────────────────────┤
│  ③ Power Query 層 : M言語の式文字列         │ PDF読込・表検出・変換
├───────────────────────────────────────────────┤
│  ④ VBA / COM 層 : PDF2ExcelMacros.bas        │ xlsx出力・後処理
└───────────────────────────────────────────────┘
```

---

## フェーズ一覧

全処理を **7 フェーズ・約 30 ステップ** に分解します。

| フェーズ | 概要 | 実行者 | 所要時間目安 |
|---|---|---|---|
| A. 起動と初期化 | BAT → PS へ制御を渡し、ワークスペースを整える | BAT / PS | 1-2 秒 |
| B. 入力の解決 | PDF 3 件を特定し staging フォルダへコピー | PS | 1-2 秒 |
| C. テンプレート準備 | マクロ付き .xlsm を用意し runtime へ複製 | PS / Excel COM | 2-5 秒 |
| D. Power Query 構成 | M 式をブックに注入し、PDF 読込の準備を完成 | PS / Excel COM | 1-2 秒 |
| E. PDF 読込と変換 | Power Query が PDF 3 件を読み、表を検出・変換 | Power Query (M) | 30-120 秒 |
| F. 後処理と検証 | 正規化列追加、メトリクス集計、VBA で xlsx 出力 | PS / VBA | 5-15 秒 |
| G. 片付け | 一時ファイル削除、ロック解放、COM 解放 | PS | 1-2 秒 |

---

## フェーズ A: 起動と初期化

### A-1. BAT がユーザーの入口となる

| # | 処理 | ファイル | 技術解説 |
|---|---|---|---|
| A-1 | ユーザーが `run_pdf2excel.bat` または `run_pdf2excel_v2.bat` を起動 | `run_pdf2excel.bat` / `run_pdf2excel_v2.bat` | 正式運用では `run_pdf2excel.bat` を使う。どちらも `VER2 Secure` を起動する |
| A-2 | `chcp 65001` でコンソールの文字コードを UTF-8 に設定 | 同上 | `chcp` は **C**hange **C**ode **P**age の略。65001 = UTF-8。日本語の表示崩れを防ぐ |
| A-3 | PowerShell に `-NoProfile -ExecutionPolicy RemoteSigned` で制御を渡す | 同上 | `-NoProfile` = 個人設定を読み込まない（環境差異を排除）。正式導線では `RemoteSigned` を前提にする |

> **専門知識: なぜ BAT と PS を分離するか？**
>
> - BAT はダブルクリックで起動できるが、ロジックを書くには貧弱
> - PowerShell は高機能だが、直接ダブルクリックで起動できない（セキュリティ制限）
> - BAT を「ランチャー」、PS を「本体」として役割分担している

### A-2. メニュースクリプト

| # | 処理 | ファイル | 技術解説 |
|---|---|---|---|
| A-4 | `run_pdf2excel_menu.ps1` が起動し、`VER2 Secure` の既定値・現在プロファイル・環境チェック導線を読み込む | `scripts/run_pdf2excel_menu.ps1` | 共通メニューはフォルダ選択ダイアログ、`[9] プロファイル選択`、`[0] 環境チェック`、完了後の `O/F` 操作を提供する |

### A-3. ワークスペース初期化

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| A-5 | `Initialize-RunWorkspace` を呼び出し | `run_pdf2excel.ps1` | 全体の前準備を一括で行う親関数 |
| A-6 | `Ensure-Workspace` で必要なフォルダ群を再帰的に作成 | `Ensure-Directory` | `VER2 Secure` では `%LOCALAPPDATA%\PDF2Excel\runtime` と `%LOCALAPPDATA%\PDF2Excel\logs`、repo 側では `output/`, `reports/`, `template/`, `config/profiles/v2/` などを整える |
| A-7 | `Rotate-LogFiles` で古いログを削除 | 同上 | 30 日以上前 or 200 件超のログを自動削除。ディスクを圧迫しない |
| A-8 | `Acquire-RunLock` で排他ロックを取得 | 同上 | **名前付き Mutex**（`Global\PDF2Excel_RunMutex`）を使用。OS レベルで同時実行を防止する |
| A-9 | `run.lock` ファイルに実行情報を JSON で書き込み | 同上 | PID、ユーザー名、開始日時を記録。ロック中に別の実行が来たとき「誰が使っているか」分かる |
| A-10 | `Compact-RuntimeArtifacts` で前回の一時ファイルを削除 | 同上 | 前回のクラッシュ残骸があれば掃除する |
| A-11 | 今回用の作業ディレクトリを作成 | 同上 | `output/runtime/runs/run_20260315_210000_123_1234/staging/` と `runtime/` |

> **専門知識: 名前付き Mutex とは？**
>
> Mutex（MUTual EXclusion = 相互排他）は、複数プロセス間で「ある資源に同時にアクセスしない」ことを保証する仕組み。`Global\` 接頭辞を付けると Windows の全セッション（ターミナルサービスを含む）で有効になる。

---

## フェーズ B: 入力の解決

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| B-1 | `Resolve-ExecutionPlan` を呼び出し | `run_pdf2excel.ps1` | 入力 PDF × プロファイル × 出力先を解決してまとめる |
| B-2 | `Resolve-InputPdfFiles` で PDF 3 件のフルパスを取得 | 同上 | フォルダ指定なら `Get-ChildItem *.pdf`、ファイル指定なら個別検証。存在チェック + 拡張子チェックを実施 |
| B-3 | `Get-ProfileConfiguration` でプロファイル JSON を読み込み | 同上 | `config/profiles/v2/construction_transfer_poc.json` を解析し、30 以上のプロパティを持つオブジェクトに変換 |
| B-4 | `Confirm-Preflight` で実行前チェックを表示 | 同上 | PDF 数、出力先、プロファイル情報を表示。`-NoConfirm` でなければ Y/N 確認 |
| B-5 | `Stage-PdfFiles` で PDF 3 件を staging フォルダへコピー | `Prepare-RunInputs` | 同名チェック後、`output/runtime/runs/.../staging/` にコピー。元ファイルには触れない |
| B-6 | `VER2 Secure` では `input/` フォルダへ同期しない | 同上 | 正式運用では今回 PDF を `input` に複製せず、staging のみで処理する |
| B-7 | `Start-Sleep -Seconds 1` で 1 秒待機 | 同上 | Excel の PDF コネクタがファイルシステムの変更を認識するための待機時間 |

### プロファイル JSON の主要設定（PoC の場合）

| 設定キー | 値 | 意味 |
|---|---|---|
| `expectedColumns` | `30` | PDF 表の想定列数 |
| `headerRowsToSkip` | `1` | 表のヘッダー行数（スキップする行数） |
| `targetRowCount` | `5` | 想定データ行数（候補選定のスコアリングに使用） |
| `allowMoreColumns` | `false` | 想定より列が多い表を許容するか |
| `multiPageMergeMode` | `sameHeader` | 複数ページの結合方式。`sameHeader` = 同ヘッダーの表を縦連結 |
| `sourceFileColumnName` | `元ファイル名` | 出力列の先頭に付ける列名 |
| `dataColumnPrefix` | `項目` | データ列の接頭辞（`項目1`, `項目2`, ...） |
| `reviewPersonColumn` | `3` | Review 時の人名列番号 |
| `reviewSiteColumn` | `5` | Review 時の現場名列番号 |
| `reviewInTimeColumn` | `6` | Review 時の入場時刻列番号 |
| `reviewOutTimeColumn` | `7` | Review 時の退場時刻列番号 |
| `normalizedTimeColumns` | `[6→正規化入場1, 7→正規化退場1, 9→正規化入場2, 10→正規化退場2]` | 時刻正規化対象列 |

---

## フェーズ C: テンプレート準備

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| C-1 | `Ensure-Template` でテンプレート xlsm の存在を確認 | `run_pdf2excel.ps1` | 初回は `build_excel_template.ps1` を子プロセスで起動して生成 |
| C-2 | テンプレートスクリプトが Excel COM で新規ブックを作成 | `build_excel_template.ps1` | `New-Object -ComObject Excel.Application` で Excel プロセスを起動 |
| C-3 | V2 テンプレートとして 5 シートを作成 | 同上 | Control, Result, Errors, Summary, Review の 5 シート |
| C-4 | 各シートのレイアウト（列幅・見出し・色）を設定 | 同上 | `Set-ControlSheetLayout` 等で事前レイアウトを描画。薄青のヘッダー色 = `15773696`（RGB の 10 進数表現） |
| C-5 | VBA マクロモジュール（`PDF2ExcelMacros.bas`）をブックに取り込み | `Import-VbaModule` | `VBProject.VBComponents.Add(1)` で標準モジュールを追加し、コードを流し込む |
| C-6 | `.xlsm` 形式（ファイル形式 52）で保存 | 同上 | xlsm = マクロ有効ブック。xlsx（51）ではマクロを保存できない |
| C-7 | `Copy-TemplateToRuntime` でテンプレートを runtime フォルダにコピー | `run_pdf2excel.ps1` | テンプレート原本は汚さず、コピーに対して作業する |

### V2 テンプレートのシート構成

| シート名 | 役割 | V1 にあるか |
|---|---|---|
| Control | メタデータ（入出力先、実行状態、プロファイル情報） | ✅ |
| Result | PDF から抽出した raw データ | ✅ |
| Errors | 読取失敗した PDF とエラー詳細 | ✅ |
| Summary | 実行結果のサマリー（PDF 別内訳、エラー分類） | ✅ |
| Review | 確認が必要な行（V2 専用） | ❌ V2 のみ |

> **専門知識: COM オートメーションとは？**
>
> COM (Component Object Model) は Windows の技術で、あるアプリケーション（PowerShell）から別のアプリケーション（Excel）をプログラムで操作する仕組み。`New-Object -ComObject Excel.Application` は「Excelを裏で起動して、PowerShell からリモコンのように操作する」という意味。

---

## フェーズ D: Power Query 構成

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| D-1 | `Open-ExcelRuntimeContext` で runtime xlsm を Excel COM で開く | `run_pdf2excel.ps1` | `$excel.Visible = $false` で見えない Excel ウィンドウを起動 |
| D-2 | Control シートに入出力先やプロファイル情報を書き込み | `Set-ControlValues` | B2=入力フォルダ, B3=出力先, B4=ログパス, B13=プロファイル名 等 |
| D-3 | Summary シートの静的セル（見出し）を設定 | `Initialize-SummarySheet` | 集計表の骨組みを事前に描画 |
| D-4 | `Configure-WorkbookQueries` で 6 つの Power Query をブックに注入 | 同上 | ここが核心。PowerShell が M 式の「文字列」を動的に生成し、ブックの Queries コレクションに設定する |

### 注入される 6 つの Power Query

| # | クエリ名 | M 式生成関数 | 役割 |
|---|---|---|---|
| 1 | `PDF2Excel_Staging` | `Get-StagingQueryFormulaV2Simple` | **最重要**。PDF 3 件を読み、表を検出・スコアリング・マージし、結果とレビュー情報を返す |
| 2 | `PDF2Excel_Result` | `Get-ResultQueryFormula` | Staging の成功行から Data 列を展開し、Result シート向けの表を作る |
| 3 | `PDF2Excel_Errors` | `Get-ErrorsQueryFormula` | Staging の失敗行を抽出し、エラー詳細を表にする |
| 4 | `PDF2Excel_FileSummary` | `Get-FileSummaryQueryFormulaV2` | PDF 別の成否・行数・Review 件数をまとめる |
| 5 | `PDF2Excel_ErrorSummary` | `Get-ErrorSummaryQueryFormula` | エラーをカテゴリ・コード別に集計する |
| 6 | `PDF2Excel_Review` | `Get-ReviewQueryFormulaV2` | Staging の Review テーブルを展開し、レビュー対象行を一覧化する |

### クエリ間の依存関係

```
PDF2Excel_Staging ←── PDF 3 件（Folder.Files）
    ├── PDF2Excel_Result       ← Staging の Data 列を展開
    ├── PDF2Excel_Errors       ← Staging の IsError=true を抽出
    ├── PDF2Excel_FileSummary  ← Staging の全行のメタ情報を集計
    ├── PDF2Excel_ErrorSummary ← Staging の ErrorCategory を集計
    └── PDF2Excel_Review       ← Staging の Review 列を展開
```

> **専門知識: Power Query (M言語) とは？**
>
> Excel に内蔵されたデータ取得・変換エンジン。M（Mashup の M）という関数型言語で式を書く。特徴:
> - **遅延評価**: 式を書いた時点では実行されず、Refresh 時に初めて動く
> - **Pdf.Tables()**: Excel 2019+ / M365 で使える関数。PDF 内の表を自動認識する
> - **Folder.Files()**: フォルダ内のファイル一覧を取得する関数。staging フォルダの全 PDF を列挙する

---

## フェーズ E: PDF 読込と変換 (Power Query の内部動作)

ここが全処理の核心です。**PDF 3 件それぞれについて**以下が実行されます。

### E-1. PDF からの表候補抽出

| # | 処理 | M 式の該当部分 | 技術解説 |
|---|---|---|---|
| E-1a | `Folder.Files("...staging")` で staging フォルダの PDF 3 件を列挙 | `Source = Folder.Files(...)` | 各行に Name, Content (バイナリ) 等が入ったテーブルを返す |
| E-1b | 拡張子 `.pdf` のファイルだけをフィルタ | `PdfFiles = Table.SelectRows(...)` | 大文字小文字を無視して `.pdf` のみ選択 |
| E-1c | 各 PDF に対し `Pdf.Tables()` を **2 通り**で呼び出す | `mergedTry` / `pageTry` | `MultiPageTables = true`（ページ結合あり）と `false`（ページ個別）の両方で試行 |

> **専門知識: Pdf.Tables() の動作原理**
>
> - PDF 内のテキスト座標（x, y）を解析し、罫線や空白の並びから「表」を推定する
> - 1 つの PDF から複数の表候補（Table001, Table002, …）が返される
> - `MultiPageTables = true` は複数ページにまたがる表を 1 つに結合しようとする
> - `MultiPageTables = false` は各ページの表を別々に返す
> - どちらが正しいかは PDF の構造による → 両方試して良い方を選ぶ設計

### E-2. 表候補のスコアリング

各候補に対してスコアを計算し、最も適した表を選びます。

| # | 処理 | 技術解説 |
|---|---|---|
| E-2a | 各候補のメタ情報を抽出 | `columnCount`（列数）, `rowCount`（行数）, `tableKind`（表種別）, `tableName`（名前） |
| E-2b | ヘッダー署名を計算 | ヘッダー行のセル値を `|` 区切りで連結した文字列。同じ構造の表を同定するために使う |
| E-2c | ページ番号を推定 | テーブル ID や名前から `PAGE1`, `PAGE2` 等のトークンを抽出 |
| E-2d | スコアを計算 | 下記の計算式参照 |

### スコア計算式

```
score = |列数 - 想定列数| × 100000
      - データ行数 × 10             ← 行が多いほど良い（負方向＝有利）
      + kindBonus                    ← preferredTableKinds にマッチで -250
      + nameBonus                    ← preferredTableNameContains にマッチで -120
      + idBonus                      ← preferredTableIdContains にマッチで -120
```

| スコア要素 | 意味 | 重み |
|---|---|---|
| 列数の差 | 想定列数（30）に近いほど有利 | × 100,000（最重要） |
| データ行数 | 行が多いほど有利（実データが多い表を優先） | × 10 |
| 表種別ボーナス | `Table` 型にマッチすると有利 | -250 |
| 名前ボーナス | `勤怠`, `現場`, `入退場` を含む表名は有利 | -120 |

> **スコアが低いほど良い候補** です。列数の差が最優先であり、30 列の表があれば他より圧倒的に有利。

### E-3. 表のマージ戦略決定

PDF 3 件それぞれについて、以下の 3 つのマージ戦略を試み、最も適したものを選びます。

| # | 戦略 | 条件 | 具体例 |
|---|---|---|---|
| E-3a | **縦連結** (Vertical Merge) | 同一列数 × 同一行数 × 同一 TableKind の表が複数ある | 6 ページ PDF: 30列×5行の表が 6 個 → 30 行に連結 |
| E-3b | **横結合** (Horizontal Merge) | 列数合計が想定列数に達する部分表群がある | 単票 PoC: 19列 + 11列 → 30 列に復元 |
| E-3c | **単一表選択** (Single) | 想定列数に一致する表が 1 つだけ | 2 ページ PDF: 30列の表が 2 個 → 先頭だけ or 縦連結 |

### 優先順位

```
横結合の結果 → 縦連結の結果 → MultiPageTables=true の最良候補 → ページ別の最良候補
```

### E-4. データ列の正規化と NormalizeText

選ばれた表に対して列名を統一的に変換します。

| # | 処理 | 前 → 後 | 技術解説 |
|---|---|---|---|
| E-4a | 元の列名を連番接頭辞に付け替え | `Column1, Column2, ...` → `項目1, 項目2, ...` | `DataColumnPrefix` + 連番 |
| E-4b | 想定列数に足りない場合は null 列を追加 | 25 列 → 30 列（項目26〜30 が null） | パディング |
| E-4c | **NormalizeText** でセル値のテキスト正規化 | `田中　太郎\n現場A` → `田中 太郎 現場A` | 改行→空白、全角空白→半角空白、連続空白→1 つ、トリム |
| E-4d | `__PageNumber` 列を追加 | 各行にページ番号を付与 | マージ後にどのページ由来かを追跡するための内部列 |

### NormalizeText の変換ルール

```
入力                     → 出力
"田中　太郎"             → "田中 太郎"       (全角空白→半角)
"田中\r\n太郎"           → "田中 太郎"       (改行→空白)
"  田中   太郎  "        → "田中 太郎"       (連続空白圧縮+トリム)
""                       → ""               (空はそのまま)
```

### E-5. Review テーブルの生成

各 PDF のデータ行を 1 行ずつ検査し、確認が必要な行を抽出します。

| # | チェック項目 | 条件 | 生成される理由文字列 |
|---|---|---|---|
| E-5a | 時刻の片方が空 | `InRaw` または `OutRaw` の長さが 0 | `入退場時刻の片方が空です` |
| E-5b | セル内に複数時刻 | `CountTimeLikeTokens` が 2 以上 | `同一セルに複数の時刻らしき文字列があります` |
| E-5c | ヘッダー差異メモ | 同一 PDF 内でヘッダー不一致の候補表がある | `同一 PDF 内にヘッダー不一致または列ずれの候補表がありました。` |

複数理由がある場合は ` / ` で連結（例: `入退場時刻の片方が空です / 同一セルに複数の時刻らしき文字列があります`）。

### E-6. Staging テーブルの最終出力

1 件の PDF につき 1 行が Staging テーブルに追加されます。

| 列名 | 内容 | PDF①の例 |
|---|---|---|
| `Name` | ファイル名 | `現場A_2026年03月.pdf` |
| `IsError` | エラーかどうか | `false` |
| `ErrorCode` | エラーコード | `null` |
| `ErrorCategory` | エラーカテゴリ | `null` |
| `UserMessage` | ユーザー向けメッセージ | `null` |
| `CandidateColumns` | 採用した表の列数 | `30` |
| `CandidateRows` | 採用した表のデータ行数 | `5` |
| `OutputRowCount` | 出力行数 | `5` |
| `ReviewCount` | Review 対象行数 | `2` |
| `PageCount` | マージしたページ数 | `1` |
| `Data` | **展開可能なテーブル** | (30列×5行のテーブル型) |
| `Review` | **展開可能なテーブル** | (7列×2行のテーブル型) |

> **専門知識: M のネストされたテーブル**
>
> M ではテーブルの「セル」にさらにテーブルを格納できる（ネスト）。`Data` 列には「この PDF から抽出した表」がまるごと入っている。`Table.ExpandTableColumn` で平たい行に展開する。

---

## フェーズ E-7: 下流クエリの動作

Staging テーブルが完成した後、下流の 5 クエリがそれを参照して各シート向けのデータを生成します。

| クエリ → シート | 抽出ロジック | PDF 3 件の場合の出力行数 |
|---|---|---|
| `PDF2Excel_Result` → Result | `IsError <> true` の行の `Data` 列を展開 | 3 件 × 5 行 = 15 行（成功時） |
| `PDF2Excel_Errors` → Errors | `IsError = true` の行を選択 | 0 行（全件成功の場合） |
| `PDF2Excel_Review` → Review | `IsError <> true` の行の `Review` 列を展開 | Review 対象行数の合計 |
| `PDF2Excel_FileSummary` → Summary | 全行のメタ情報（成否、行数、Review 数）を集計 | 3 行 |
| `PDF2Excel_ErrorSummary` → Summary | エラーをカテゴリ×コードで集計 | 0 行（全件成功の場合） |

---

## フェーズ F: 後処理と出力

### F-1. シートへの読込

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| F-1 | Result クエリ → Result シート `A1` へ読み込み | `Load-WorkbookQueryToWorksheet` | OLEDB 接続（`Microsoft.Mashup.OleDb.1`）でクエリ結果をシート上のテーブル（ListObject）に展開する |
| F-2 | Errors クエリ → Errors シート `A1` へ | 同上 | 同様 |
| F-3 | Review クエリ → Review シート `A1` へ | 同上 | V2 のみ |
| F-4 | FileSummary クエリ → Summary シート `A14` へ | 同上 | 既存のサマリーメトリクスの下に表を配置 |
| F-5 | ErrorSummary クエリ → Summary シート `S14` へ | 同上 | V2 は列が多いため右寄り（S14）に配置 |

> **専門知識: OLEDB プロバイダとは？**
>
> `Microsoft.Mashup.OleDb.1` は Power Query の結果を「データベースのテーブル」として Excel に公開するドライバ。Excel 内部で `SELECT * FROM [PDF2Excel_Result]` というような SQL に近い形でデータを取得している。

### F-2. V2 専用の後処理（正規化列追加）

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| F-6 | `Apply-VersionSpecificWorkbookEnrichments` を呼び出し | `run_pdf2excel.ps1` | V1 のときは即 return。V2 のときだけ以下を実行 |
| F-7 | `Add-V2ResultNormalizedColumns` で Result シートに正規化列を追加 | 同上 | PowerShell が Excel COM でセルを 1 行ずつ読み、`Normalize-TimeText` で変換し、結果を書き戻す |
| F-8 | `Add-V2ReviewNormalizedColumns` で Review シートに正規化列を追加 | 同上 | Review シートの `InRaw` / `OutRaw` を同様に正規化 |

### Normalize-TimeText の動作

1 つの時刻セルに対して以下のパイプラインが走ります。

| ステップ | 入力 | 出力 | 説明 |
|---|---|---|---|
| 1. null チェック | `null` | `Status=EMPTY` | null は空として処理 |
| 2. 数値チェック | `0.385417` | `09:15 / 555分` | 0〜1 の小数は Excel の時刻比率として解釈（× 1440 で分換算） |
| 3. 全角→半角 | `０８：００` | `08:00` | 全角数字・コロンを半角化 |
| 4. 小数再チェック | `0.75` | `18:00 / 1080分` | 半角化後にもう一度小数チェック |
| 5. 日本語変換 | `9時 15分` | `9:15` | `時`→`:`, `分`→削除, `頃`→削除, 空白除去 |
| 6. 不完全補完 | `18:` or `18` | `18:00` | 分が省略されたケースを補完 |
| 7. 範囲検証 | `25:00` | `Status=INVALID` | 0-23 時、0-59 分の範囲外はエラー |
| 8. 最終出力 | `9:15` | `09:15 / 555分 / OK` | 正規化テキスト + 分換算 + ステータス |

### Result シートに追加される列（PoC の場合）

| 列名 | 由来 | 例 |
|---|---|---|
| `正規化入場1` | 項目6 を正規化した HH:mm | `08:00` |
| `正規化入場1_分` | 0:00 からの分換算 | `480` |
| `正規化退場1` | 項目7 を正規化した HH:mm | `17:30` |
| `正規化退場1_分` | 分換算 | `1050` |
| `正規化入場2` | 項目9 を正規化 | `09:15` |
| `正規化入場2_分` | 分換算 | `555` |
| `正規化退場2` | 項目10 を正規化 | `18:00` |
| `正規化退場2_分` | 分換算 | `1080` |
| `時刻正規化状態` | 全列の総合ステータス | `OK` or `要確認` |
| `時刻確認メモ` | 問題があった場合の詳細 | `正規化入場2: 時刻として解釈できません。` |

### F-3. メトリクス集計と出力

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| F-9 | `Measure-WorkbookOutcome` で各シートの行数を集計 | `run_pdf2excel.ps1` | ListObject.DataBodyRange.Rows.Count で実データ行数を取得 |
| F-10 | Control シートにメトリクスを書き込み | `Set-ControlMetrics` | B7=PDF数, B8=Result行数, B9=Error件数, B10=成功数, B11=失敗数, B12=処理秒数 |
| F-11 | Summary シートにサマリーを書き込み | `Set-SummaryMetrics` | PDF 数、成否、Review 件数、処理時間、プロファイル名 |
| F-12 | Control の B6 を `出力準備完了` に更新 | `Update-WorkbookSummaryState` | 状態遷移: 待機中 → 準備完了 → 出力準備完了 |

### F-4. xlsx として出力

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| F-13 | `Publish-WorkbookOutput` を呼び出し | `run_pdf2excel.ps1` | まず VBA マクロでの出力を試み、失敗したら PowerShell フォールバック |
| F-14 | (成功時) VBA マクロ `ExportResultAsXlsx` を実行 | `PDF2ExcelMacros.bas` | xlsm から必要なシートだけを新しいブックにコピーし、xlsx（ファイル形式 51）で保存 |
| F-15 | (VBA 失敗時) `Export-WorkbookDirectly` で直接保存 | `run_pdf2excel.ps1` | `Workbook.SaveAs(path, 51)` で xlsx として上書き保存 |

> **専門知識: なぜ VBA と PS のフォールバックがあるか？**
>
> - VBA マクロの実行には「マクロの実行を許可する」セキュリティ設定が必要
> - 企業環境ではマクロが無効化されていることがある
> - その場合でも PowerShell の COM 操作で同等の出力ができるため、フォールバックを用意

---

## フェーズ G: 片付け

| # | 処理 | 関数 | 技術解説 |
|---|---|---|---|
| G-1 | ランタイム Excel ブックを閉じる | `Close-ExcelRuntimeContext` | `Workbook.Close($false)` = 保存せず閉じる（xlsm は一時ファイルであるため） |
| G-2 | Excel プロセスを終了 | 同上 | `Excel.Quit()` で COM 経由で終了。これを忘れると EXCEL.exe がゾンビとして残る |
| G-3 | COM オブジェクトを解放 | `Release-ComObject` | `Marshal.FinalReleaseComObject()` で参照カウントをゼロにする |
| G-4 | 作業ディレクトリを削除 | `Finalize-RunWorkspace` | `output/runtime/runs/run_...` を丸ごと削除 |
| G-5 | Mutex を解放 | `Release-RunLock` | 名前付き Mutex を `ReleaseMutex()` + `Dispose()` で開放 |
| G-6 | `run.lock` ファイルを削除 | 同上 | ロック情報ファイルを消す |
| G-7 | ガベージコレクションを強制実行 | メインスクリプト末尾 | `[GC]::Collect()` + `WaitForPendingFinalizers()` で COM の残骸を確実に掃除 |

> **専門知識: COM オブジェクトのメモリリーク**
>
> PowerShell から Excel COM を使うとき、最も多いバグは「COM オブジェクトの解放漏れ」。解放しないと Excel.exe がバックグラウンドに残り続ける。対策:
> - 使い終わった COM オブジェクトは必ず `Release-ComObject` を呼ぶ
> - `try/finally` で確実に実行する
> - 最後に `[GC]::Collect()` でガベージコレクションを強制する

---

## 最終出力物

PDF 3 件が正常に処理された場合、以下の xlsx が生成されます。

### 出力ブックの構造

| シート | 内容 | PDF 3 件 × 5 行の場合 |
|---|---|---|
| Control | 実行メタデータ | 18 行（固定） |
| Result | raw データ + 正規化列 | ヘッダー 1 行 + データ 15 行 = 16 行、元の 30 列 + 正規化 10 列 = **40 列** |
| Errors | エラー一覧 | ヘッダー 1 行のみ（全件成功時） |
| Summary | 全体サマリー + PDF 別内訳 + エラー分類 | メトリクス + 3 行の PDF 別テーブル |
| Review | 確認要の行 | 問題のある行数に応じて可変 |

### Result シートの列構成（40 列）

| 列グループ | 列数 | 列名例 |
|---|---|---|
| 元ファイル名 | 1 | `元ファイル名` |
| raw データ列 | 30 | `項目1` 〜 `項目30` |
| 正規化入退場 | 4 | `正規化入場1`, `正規化退場1`, `正規化入場2`, `正規化退場2` |
| 分換算 | 4 | `正規化入場1_分`, `正規化退場1_分`, `正規化入場2_分`, `正規化退場2_分` |
| 時刻状態 | 1 | `時刻正規化状態` |

---

## 処理時間の配分（PDF 3 件の場合の目安）

| フェーズ | 処理時間 | 全体に占める割合 |
|---|---|---|
| A. 起動と初期化 | ~2 秒 | 1% |
| B. 入力の解決 | ~2 秒 | 1% |
| C. テンプレート準備 | ~5 秒（初回） / 0 秒（2 回目以降） | 3% |
| D. Power Query 構成 | ~2 秒 | 1% |
| **E. PDF 読込と変換** | **~120 秒** | **80%** |
| F. 後処理と出力 | ~15 秒 | 10% |
| G. 片付け | ~2 秒 | 1% |
| **合計** | **~150 秒** | |

> 処理時間の大半は Power Query による PDF 解析（`Pdf.Tables()`）に費やされます。PDF のページ数やレイアウトの複雑さに比例します。

---

## 用語集

| 用語 | 説明 |
|---|---|
| **BAT** | Windows バッチファイル。ダブルクリックで実行できるスクリプト |
| **COM** | Component Object Model。Windows アプリ間の通信規格 |
| **Excel COM** | PowerShell から Excel を操作する技術。`New-Object -ComObject Excel.Application` |
| **M 言語** | Power Query のプログラミング言語。関数型で遅延評価 |
| **Mutex** | 排他制御の仕組み。同時実行を防止する |
| **OLEDB** | データベースアクセスの標準規格。Power Query の結果を Excel に渡す際に使用 |
| **Power Query** | Excel 組み込みのデータ取得・変換エンジン |
| **Pdf.Tables()** | Power Query の関数。PDF 内の表を認識して M のテーブルとして返す |
| **staging** | 一時的な作業領域。元ファイルを汚さずに処理するためにコピーする場所 |
| **xlsm** | マクロ有効ブック。VBA コードを含む Excel ファイル形式 |
| **xlsx** | マクロなしブック。最終出力はこの形式 |
| **VBA** | Visual Basic for Applications。Excel のマクロ言語 |
| **ListObject** | Excel のテーブル（構造化参照）のオブジェクト。ヘッダー + データ行で構成 |
| **正規化** | 表記のゆれを統一する処理。`08：00` → `08:00` など |
