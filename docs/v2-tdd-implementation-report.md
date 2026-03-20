# V2 TDD 実装報告

- 作成日: 2026-03-15 15:35 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## 注記

- 本文は TDD 導入時点の実装記録です。現行の起動導線、環境チェック、プロファイル設定ウィザードの仕様は [README.md](../README.md) と [user-manual.md](user-manual.md) を優先してください。

## 1. 概要

今回の対応では、`VER1` を壊さないことを最優先にしつつ、`VER2` に対して次を実装しました。

- 特定レイアウトの PoC 帳票の横分割ページを 1 行 30 列へ復元する V2 専用マージ
- `2ページ同一列` / `6ページ同一列` を安全に縦連結する V2 専用マージ
- 生データ列を保持したまま、`正規化入場*` / `正規化退場*` / `*_分` / `時刻正規化状態` / `時刻確認メモ` を追加
- `08：00`、`9時 15分`、`18時`、Excel 時刻比率文字列を正規化する共通関数
- V2 単票、2ページ、6ページを含むテスト強化
- handoff 配布物の再生成

## 2. 進め方

### 2-1. 先に現状をコミット

作業開始時点の全状態を先にコミットしました。

- コミット: `85802ab`
- メッセージ: `chore: snapshot current v1 v2 split state`

これにより、以降の V2 変更を安全に追えるようにしました。

### 2-2. TDD で赤から開始

まずテストを先に足しました。

- ユニットテスト
  - `Normalize-TimeText` に対して `08：00`、`9時 15分`、`18時`、Excel 時刻比率文字列のケースを追加
  - V2 プロファイルに `normalizedTimeColumns` が載ることを追加
  - V2 用の出力列名に正規化列が含まれることを追加
  - V2 用クエリ文字列に `GetHeaderSignature`、`__PageNumber`、ヘッダー差異メモ文言が含まれることを追加
- 統合テスト
  - V2 PoC ケースで `Result` / `Review` に正規化列が出ることを追加
  - `9時15分 -> 555分`、`18時00分 -> 1080分` を確認する検証を追加

この段階では、想定どおりユニットテストは失敗しました。

## 3. 実装内容

### 3-1. 時刻正規化関数

共通関数 [pdf2excel.common.ps1](/C:/Work_Codex/PDF2Excel/scripts/pdf2excel.common.ps1) に次を追加しました。

- `Normalize-TimeText`
- `Convert-MinutesToTimeText`
- `Get-NormalizedTimeColumnDefinitions`
- `Get-ResultOutputColumnNames`

`Normalize-TimeText` では次を吸収します。

- 全角コロン
- `時` `分` 表記
- 余計な空白
- `18時` のような分省略
- `0.385416666...` のような Excel 時刻比率文字列

出力は次の情報を返します。

- 正規化文字列 `HH:mm`
- 0:00 からの分換算
- 状態 `OK / EMPTY / INVALID`
- 補足メモ

### 3-2. V2 プロファイル

[construction_transfer_poc.json](/C:/Work_Codex/PDF2Excel/config/profiles/v2/construction_transfer_poc.json) に `normalizedTimeColumns` を追加しました。

- `項目6 -> 正規化入場1`
- `項目7 -> 正規化退場1`
- `項目9 -> 正規化入場2`
- `項目10 -> 正規化退場2`

これにより、V2 側で「どの生データ列を時刻正規化対象にするか」を明示できるようにしました。

### 3-3. V2 の表マージ

[run_pdf2excel.ps1](/C:/Work_Codex/PDF2Excel/scripts/run_pdf2excel.ps1) の V2 クエリ生成では、`Pdf.Tables` の実態に合わせて 2 系統のマージを実装しました。

1. 横結合
   - 単票 PoC は `19列 + 11列` に分かれていたため、同じ行数の部分表を横方向に復元して 30 列へ統合
2. 縦結合
   - `2ページ同一列` / `6ページ同一列` は 30 列表が複数ページに再出現するため、同列構成の候補を縦連結

ヘッダー署名そのものも取得しますが、6ページケースでは月ごとの見出し差異が混じるため、実結合条件は `列数 + 行数 + TableKind` を主に使い、ヘッダー差異は `Review` 側メモへ送る構成にしました。

### 3-4. Workbook 後処理

Power Query では生データ転記を優先し、正規化列は Excel 読込後に PowerShell 側で追加しています。

追加した後処理は次です。

- `Add-V2ResultNormalizedColumns`
- `Add-V2ReviewNormalizedColumns`
- `Apply-VersionSpecificWorkbookEnrichments`

役割は次の通りです。

- `Result`
  - 生データを保持したまま正規化列を末尾に追加
- `Review`
  - 確認対象行に対して正規化結果も表示

分析向けには `*_分` 列を持たせ、表示向けには `HH:mm` 文字列列を持たせています。

## 4. 調査で分かったこと

今回の実装で重要だったのは、PoC PDF の構造を決め打ちで想像しないことでした。  
一時的に `Pdf.Tables` を直接調べた結果、次が分かりました。

- 単票 PoC
  - `Table001 = 19列`
  - `Table002 = 11列`
  - つまり 30 列 1 表ではなく、横分割ページ
- `2ページ同一列`
  - `30列 x 2 表`
- `6ページ同一列`
  - `30列 x 6 表`

この差を吸収しないと、単票は `TABLE_NOT_FOUND`、6ページは 2 ページ分しか連結しない、という不具合になります。

## 5. テスト強化

### 5-1. ユニットテスト

対象:

- 共通関数
- V2 プロファイル解釈
- V2 クエリ生成文字列

追加した主な観点:

- `08：00 -> 08:00 / 480`
- `9時 15分 -> 09:15 / 555`
- `18時 -> 18:00 / 1080`
- `0.385416666666667 -> 09:15 / 555`
- V2 出力列に正規化列が出ること
- V2 クエリにページ番号とヘッダー差異メモが含まれること

### 5-2. 統合テスト

主に次を確認しました。

- V1 正常系が引き続き成功すること
- V2 単票 PoC が 6 行で出ること
- V2 2ページ同一列が 11 行で出ること
- V2 6ページ同一列が 31 行で出ること
- `Result` / `Review` に正規化列が出ること
- `555` 分、`1080` 分が出ること

## 6. 実施ログ

### 6-1. テスト基盤で直したもの

フル統合スイートを再実行した際、PowerShell 5.1 で `Stop-Job -Force` が無効なため、テストランナーが途中で落ちる問題がありました。  
[run_integration_tests.ps1](/C:/Work_Codex/PDF2Excel/tests/run_integration_tests.ps1) を修正し、`Stop-Job` 互換に直しました。

### 6-2. 代表検証

実施日時:

- 2026-03-15 15:19:32 JST - 2026-03-15 15:29:34 JST

対象環境:

- Windows
- PowerShell 5.1
- Excel(M365) COM

対象機能:

- 共通時刻正規化
- V1 PowerShell 正常変換
- V2 単票 PoC
- V2 2ページ同一列
- V2 6ページ同一列

実行シナリオ:

- `tests/run_unit_tests.ps1`
- `tests/run_integration_tests.ps1 -CaseName 'PowerShell 経由の正常変換'`
- `tests/run_integration_tests.ps1 -CaseName '建設現場転記PoCの変換'`
- `tests/run_integration_tests.ps1 -CaseName 'V2 2ページ同一列の変換'`
- `tests/run_integration_tests.ps1 -CaseName 'V2 6ページ同一列の変換'`

結果概要:

- Unit: 成功
- V1 PowerShell: 成功
- V2 PoC: 成功
- V2 2page: 成功
- V2 6page: 成功

所要時間または主要な応答時間:

- Unit: 2 秒
- V1 PowerShell: 134 秒
- V2 PoC: 154 秒
- V2 2page: 152 秒
- V2 6page: 161 秒

エラー有無:

- 最終代表検証ではなし

失敗時の原因推定:

- 実装中には次を修正済み
  - 横分割ページを 30 列表と見なしていた誤り
  - 6ページケースでヘッダー差異を厳密比較しすぎていた誤り
  - Excel COM への型代入と列フォーマット指定の不安定さ
  - テストランナーの `Stop-Job -Force` 非互換

## 7. 文書・配布物

更新または再生成したもの:

- [README.md](/C:/Work_Codex/PDF2Excel/README.md)
- [user-manual.md](/C:/Work_Codex/PDF2Excel/docs/user-manual.md)
- [v2-data-transfer-proposal.md](/C:/Work_Codex/PDF2Excel/docs/v2-data-transfer-proposal.md)
- `handoff/generated/PDF2Excel_V1_Minimal`
- `handoff/generated/PDF2Excel_V2_Minimal`
- 各 zip 配布物

## 8. 残課題

今回で「生データを壊さず、時刻を分析しやすくする」基盤までは入りました。  
まだ残るのは次です。

- 氏名・現場名の辞書正規化
- Review の理由分類の細分化
- 正規化列を使った縦持ち集計シート
- 統合スイート全件の定常化

ただし、今回の目的だった `VER2` の生データ転記安定化、多ページマージ、時刻正規化、V1 非破壊は満たしています。
