# V2 残課題 実装計画

- 作成日: 2026-03-15 20:58 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20
- ステータス: **承認済み・未着手**

## 注記

- 本文は当時の残課題計画です。現状の完了状況は [CHANGELOG.md](../../CHANGELOG.md) と [technical-description.md](../01-current/04-technical-description.md) を優先してください。

## 1. 概要

V2 TDD 実装報告 ([v2-tdd-implementation-report.md](/C:/Work_Codex/PDF2Excel/docs/04-reports/01-v2-tdd-implementation-report.md)) で挙げた残課題 4 件を実装する計画です。

**設計方針**: 特定帳票向けの PoC を前提に実装しますが、プロファイル JSON の設定だけで任意の帳票に適用できる汎用設計とします。ハードコードされた列番号や帳票固有のロジックは一切入れません。

---

## 2. 変更①: 空白位置の正規化

### 目的

PDF 抽出時に混入する全角空白・連続空白の表記揺れを吸収し、氏名や現場名を統一された形式にする。

### 汎用設計

プロファイル JSON に `textNormalizeColumns` 配列を追加。どの列を正規化するか、出力列名は何にするかをプロファイルで定義する。

```json
{
  "textNormalizeColumns": [
    { "sourceColumn": 3, "displayName": "正規化氏名" },
    { "sourceColumn": 5, "displayName": "正規化現場名" }
  ]
}
```

- 列番号もプロファイルから読み込むため、勤怠表以外の帳票でも任意の列に適用可能
- `textNormalizeColumns` を省略した場合、この機能は無効（既存プロファイルへの影響なし）

### 対象ファイル

| ファイル | 変更内容 |
|---|---|
| `scripts/pdf2excel.common.ps1` | `Normalize-WhitespaceText` 関数を追加。全角空白→半角化、連続空白→1 つに圧縮、前後トリム。`Get-TextNormalizeColumnDefinitions` 関数を追加（プロファイルから定義を読み込み） |
| `scripts/run_pdf2excel.ps1` | `Get-ProfileConfiguration` で `textNormalizeColumns` を読み込み。`Add-V2ResultNormalizedColumns` の末尾でテキスト正規化列を Result シートに追加。Review シートの `PersonRaw` / `SiteRaw` 隣に正規化値を表示 |
| `config/profiles/v2/construction_transfer_poc.json` | `textNormalizeColumns` を追加（列 3=氏名、列 5=現場名） |

### テスト

- ユニット: `Normalize-WhitespaceText` の全角空白・連続空白・トリムテスト
- ユニット: プロファイル読み込み時に `textNormalizeColumns` が解釈されること
- 統合: V2 PoC で `正規化氏名` / `正規化現場名` が Result に出ること

---

## 3. 変更②: Review の理由分類の細分化

### 目的

Review シートのレビュー対象行に対し、なぜレビューが必要かのカテゴリコードを付け、フィルタリングや集計を容易にする。

### 汎用設計

理由の判定ルールは Power Query 内の `BuildReviewTable` 関数で実施するが、判定条件はプロファイルの `reviewMappings` に依存するため、プロファイルを変えれば異なる帳票に適用可能。新しい理由カテゴリの追加もコード側の拡張だけで対応可能。

### カテゴリコード体系

| カテゴリコード | 説明 | 対象列の参照元 |
|---|---|---|
| `TIME_MISSING` | 入退場時刻の片方が空 | `reviewInTimeColumn` / `reviewOutTimeColumn` |
| `TIME_MULTI` | 同一セルに複数の時刻文字列 | 同上 |
| `NAME_EMPTY` | 氏名または現場名が空 | `reviewPersonColumn` / `reviewSiteColumn` |
| `ROW_EMPTY` | 全データ列が空の行 | 全列 |
| `HEADER_MISMATCH` | ページ間のヘッダー差異 | PDF 構造 |

- カテゴリコードは `Reason` 文字列の先頭に `[CODE]` 形式で付与
- 複数理由がある場合は `Reason` を ` / ` で連結、`ReasonCategory` 列にはコードだけをカンマ区切りで保持
- `reviewPersonColumn` 等が `null` の場合、対応する判定はスキップ（汎用性を維持）

### 対象ファイル

| ファイル | 変更内容 |
|---|---|
| `scripts/run_pdf2excel.ps1` | `BuildReviewTable`（V2/V2Simple 両方）の理由生成にカテゴリコードを付与。`NAME_EMPTY` / `ROW_EMPTY` の判定を追加。`Get-ReviewQueryFormulaV2` に `ReasonCategory` 列を追加。`Add-V2ReviewNormalizedColumns` で `ReasonCategory` 列を書き込み |

### テスト

- ユニット: V2 クエリ文字列に `TIME_MISSING` / `ReasonCategory` が含まれること
- 統合: V2 PoC で Review に `ReasonCategory` 列が存在すること

---

## 4. 変更③: 正規化列を使った縦持ち集計シート

### 目的

横持ち（1行に入場1・退場1・入場2・退場2…）のデータを、1レコード＝1回の入退場に展開した `Analysis` シートを自動生成し、ピボットテーブルやグラフでの分析を容易にする。

### 汎用設計

縦持ち展開の対象列はプロファイルの `normalizedTimeColumns` を利用する。この配列が空のプロファイル（V1 や時刻列のない V2）では `Analysis` シートは生成しない。

入退場のペアリングもプロファイルの `analysisTimeGroups` で定義可能にする:

```json
{
  "analysisTimeGroups": [
    { "inColumn": 6, "outColumn": 7, "label": "入退場1" },
    { "inColumn": 9, "outColumn": 10, "label": "入退場2" }
  ]
}
```

- `analysisTimeGroups` が省略された場合は `normalizedTimeColumns` を 2 つずつ自動ペアリング（入場/退場の交互配置を仮定）
- ペア化できない奇数列は単独行として展開

### Analysis シートの出力列

| 列名 | 内容 |
|---|---|
| `元ファイル名` | Result の `SourceFileColumnName` の値 |
| `行番号` | Result 上の行番号（ヘッダー除き） |
| `氏名` | `reviewPersonColumn` の値（設定があれば） |
| `現場名` | `reviewSiteColumn` の値（設定があれば） |
| `入退場グループ` | `label` の値（例: `入退場1`） |
| `入場時刻` | 正規化入場の `HH:mm` |
| `入場_分` | 0:00 からの分換算 |
| `退場時刻` | 正規化退場の `HH:mm` |
| `退場_分` | 0:00 からの分換算 |
| `滞在_分` | 退場_分 - 入場_分（計算可能な場合） |
| `時刻状態` | OK / EMPTY / INVALID |

### 対象ファイル

| ファイル | 変更内容 |
|---|---|
| `scripts/run_pdf2excel.ps1` | `Build-V2AnalysisSheet` 関数を新規追加。`Apply-VersionSpecificWorkbookEnrichments` から呼び出し。`Open-ExcelRuntimeContext` で `Analysis` シートを事前作成 |
| `scripts/run_pdf2excel.ps1` | `Get-ProfileConfiguration` で `analysisTimeGroups` を読み込み。省略時は `normalizedTimeColumns` から自動生成 |
| `config/profiles/v2/construction_transfer_poc.json` | `analysisTimeGroups` を追加 |

### テスト

- 統合: V2 PoC で `Analysis` シートが存在し、行数が Result の 2 倍（2 ペアのため）であること
- 統合: V2 2ページ / 6ページでも `Analysis` シートが正しく展開されること
- 統合: V1 実行で `Analysis` シートが生成されないこと

---

## 5. 変更④: 統合スイート全件の定常化

### 目的

`-CaseName` 指定なしで全件安定実行を可能にし、リグレッション検出を定常化する。

### 対象ファイル

| ファイル | 変更内容 |
|---|---|
| `tests/run_integration_tests.ps1` | フィクスチャ初期化の冪等化（既存 fixture はスキップ可能に）。50件一括のタイムアウトを 480→600 秒に拡大。`Invoke-TestProcess` に 1 回リトライ機能を追加（Excel COM の一時的失敗用）。新機能（Analysis シート、ReasonCategory 列、テキスト正規化列）の検証項目を既存ケースに追加。テスト結果レポートに全件結果を含める |

### 検証

```powershell
# CaseName なしで全件を一貫実行
powershell -NoProfile -ExecutionPolicy Bypass -File tests\run_integration_tests.ps1
```

---

## 6. 汎用性の担保ポイントまとめ

| 機能 | 汎用化の仕組み |
|---|---|
| 空白正規化 | `textNormalizeColumns` でプロファイルから対象列を指定 |
| Review 理由分類 | `reviewMappings` の null チェックで判定を自動スキップ |
| 縦持ち集計 | `analysisTimeGroups` でペアリングを定義、省略時は自動推定 |
| 縦持ち集計の有効/無効 | `normalizedTimeColumns` が空なら Analysis シートは生成しない |

新しい帳票プロファイルを追加する際は、JSON に上記の設定を追加するだけで全機能を利用可能。コード側の変更は不要。

---

## 7. 実行順序

1. ① 空白正規化（共通関数 → プロファイル → PowerShell 後処理）
2. ② Review 理由分類（Power Query → PowerShell 後処理）
3. ③ 縦持ち集計シート（PowerShell 後処理 → シート生成）
4. ④ 統合スイート定常化（テスト全件実行で①②③を検証）

各ステップ完了ごとにユニットテスト → 代表統合テストで回帰確認。

---

## 8. V1 への影響

- V1 プロファイルには `textNormalizeColumns` / `analysisTimeGroups` / `normalizedTimeColumns` が存在しないため、全機能が自動的に無効
- V1 の Power Query (`Get-StagingQueryFormula`) には `BuildReviewTable` が存在しないため、Review 理由分類も影響なし
- V1 の既存テストケースは全件パスすることを統合テストで確認
