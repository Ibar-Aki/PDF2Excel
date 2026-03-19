# V2 安定化・時刻正規化 実装報告

- 作成日: 2026-03-15 22:45 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## 注記

- 本文は安定化実装時点の報告書です。現行の正式導線は `run_pdf2excel.bat` = `VER2 Secure` であり、最新仕様は [README.md](../README.md) と [technical-description.md](technical-description.md) を優先してください。

## 1. 目的

この報告は、`VER2` を「変な PDF が混じっても誤結合しにくい生データ転記基盤」に寄せるために実施した改修内容を、調査から実装、テスト、handoff 反映まで時系列で記録したものです。

今回の対象は `VER2` のみです。`VER1` の入口、既存テンプレート、既存プロファイル、既存運用は変えない方針で進めました。

## 2. 実施前の所見

実装前に確認した主要論点は次です。

- `sameHeader` を名乗っているが、実際の縦結合条件がヘッダー一致を十分見ていない
- 横分割結合が「最初に 30 列へ届いた候補」を採用するため、異物表の混入に弱い
- `*_分` 列が文字列で書かれており、分析用の数値列として不十分
- `Review` が第 1 時刻ペア中心で、第 2 時刻ペアの異常を拾い切れていない
- `24:00` を有効値として扱えていない

## 3. 実装方針

### 3-1. V1 非破壊

- 変更対象は `scripts/run_pdf2excel.ps1` と `scripts/pdf2excel.common.ps1` の `VER2` 系分岐を中心に限定
- `config/profiles/v1`、`template` の V1 資産、V1 起動 BAT は未変更
- 代表 V1 統合テストで回帰を確認

### 3-2. TDD の進め方

先にユニットテストと V2 統合テストを増やし、失敗を観測してから実装しました。

- 時刻正規化の正常系と境界値
- `Get-StagingQueryFormula` に新ロジック断片が入ること
- V2 PoC 単票の `24:00`
- `2ページ同一列`
- `6ページ同一列`
- `ヘッダー不一致負例`
- `時刻確認負例`

## 4. 実装内容

### 4-1. 時刻正規化の強化

`scripts/pdf2excel.common.ps1` を中心に次を実装しました。

- `Normalize-TimeText` の拡張
  - `08：00`
  - `9時 15分`
  - `18時`
  - Excel 時刻比率文字列
  - `24:00`
  - `24:00:00`
- `24:30` は `INVALID`
- `*_分` は分換算の数値を返却
- `Get-TimeNormalizationAudit` を追加し、複数時刻候補、片側空、解釈不能の監査理由を統合

### 4-2. Review の動的化

`normalizedTimeColumns` を基準に、`Review` が全対象時刻列を見られるように変更しました。

- `ReviewRawColumnName` を定義
- `Get-ReviewQueryFormulaV2` を動的列生成へ変更
- `Add-V2ReviewNormalizedColumns` を profile 駆動へ変更
- `Reason` と `時刻確認メモ` を両立させる構成へ変更

### 4-3. sameHeader の厳格化

`Get-StagingQueryFormulaV2Simple` の M 式で次を実装しました。

- `CanonicalHeaderSignature` を追加
  - 空白
  - 改行
  - `ページ`
  - `管理者`
  - `2026年04月` のような月トークン
  を吸収して署名化
- `PreferredTableKinds` を実質条件にし、`Page` 候補を縦結合競合から除外
- ページ番号順の連続性確認
- 連続しない full-table 群、またはヘッダー不一致群は `TABLE_GROUP_AMBIGUOUS`

### 4-4. 横分割結合の安全化

PoC 単票では `Pdf.Tables` が `19列` と `11列` を別候補として返し、しかも `Page 1` と `Page 2` に分かれる実例がありました。  
このため、「同一物理ページ」前提の判定ではなく、次に修正しました。

- `TableKind + DataRowCount` 単位で partial 候補をグループ化
- 合計列数が `ExpectedColumns` に一致する組み合わせを探索
- 一意に組める場合のみ採用
- 複数組み合わせが成立する場合は `PARTIAL_TABLE_AMBIGUOUS`

## 5. 実装中に見つかった不具合と修正

### 5-1. M 関数名の誤り

- `Number.Min` を使っており、Power Query M で未定義
- `2ページ同一列` の時点で初めて踏み、`Result` クエリが落ちた
- `if ... then ... else ...` に置換して修正

### 5-2. ヘッダー署名にデータ行を混ぜていた

- `HeaderRowsToSkip + 2` 行を署名へ含めていたため、ページごとに署名が変わった
- `sameHeader` の縦結合が成立しなかった
- 署名対象をヘッダー行のみに修正

### 5-3. 横分割候補のページ番号誤解釈

- 単票 PoC の左右表片が `Page 1` / `Page 2` として返ることを確認
- 同一ページ判定をやめ、候補組み合わせの一意性判定へ変更

### 5-4. Excel による時刻再解釈

- `正規化退場2 = 24:00` を文字列で入れても、Excel が数値時刻として再解釈するケースがあった
- display 列を `@`、`*_分` 列を `0` に明示設定して解決

## 6. テスト追加

### 6-1. ユニットテスト

追加または強化した観点:

- `24:00`
- `24:30`
- `24:00:00`
- `ReviewRawColumnName`
- V2 M 式に `GetCanonicalHeaderSignature`
- V2 M 式に `FindHorizontalMergeSequences`
- V2 M 式に `TABLE_GROUP_AMBIGUOUS`
- V2 M 式に `正規化入場1_raw`

### 6-2. 統合テスト

追加または強化したシナリオ:

- `建設現場転記PoCの変換`
  - `24:00 -> 1440`
  - `正規化退場2 = 24:00`
  - `正規化退場2_分` が `Double`
- `V2 2ページ同一列の変換`
  - 行数
  - 氏名の重複なし
- `V2 6ページ同一列の変換`
  - 行数
  - 同一ファイルからの連続結合
- `V2 ヘッダー不一致負例の分離`
  - `Result` へ混ぜず `Errors` へ送る
- `V2 時刻確認負例の分離`
  - `24:30`
  - 片側時刻のみ
  を `Review` で拾う

## 7. 実施手順

1. 既存構成と V2 現状を確認
2. ユニットテストを追加して失敗を確認
3. 時刻正規化ロジックを更新
4. `Review` を `normalizedTimeColumns` 駆動へ変更
5. V2 ステージング M を厳格化
6. 単票 PoC を実行し、`TABLE_NOT_FOUND` を解消
7. `2ページ同一列` を実行し、M 関数誤りと署名範囲を修正
8. `6ページ同一列` を実行し、月トークン除去を追加
9. 負例サンプルと統合テストを追加
10. README / user-manual / samples / proposal / index を更新
11. handoff を再生成
12. 代表 V1 / V2 を再実行して完了判定

## 8. 実費テスト結果

- 実施日時: 2026-03-15 21:06:29 JST - 2026-03-15 22:45 JST
- 対象環境: Windows / PowerShell 5.1 / Excel(M365) COM
- 対象機能:
  - V1 標準変換
  - V2 単票 PoC
  - V2 2ページ同一列
  - V2 6ページ同一列
  - V2 ヘッダー不一致負例
  - V2 時刻確認負例
  - 共通時刻正規化
- 実行シナリオ:
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_unit_tests.ps1`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'PowerShell 経由の正常変換'`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName '建設現場転記PoCの変換'`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'V2 2ページ同一列の変換'`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'V2 6ページ同一列の変換'`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'V2 ヘッダー不一致負例の分離'`
  - `powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run_integration_tests.ps1 -CaseName 'V2 時刻確認負例の分離'`
- 結果概要:
  - ユニットテスト `27 / 27` 成功
  - 代表 V1 統合テスト成功
  - V2 正常系 3 件成功
  - V2 負例 2 件成功
- 所要時間または主要な応答時間:
  - Unit: 約 2 秒
  - V1 代表統合: 約 140 秒
  - V2 単票 PoC: 約 150 秒
  - V2 2ページ同一列: 約 170 秒
  - V2 6ページ同一列: 約 180 秒
  - V2 負例: 各 160 秒前後
- エラー有無:
  - 最終結果ではなし
  - 実装途中では `TABLE_NOT_FOUND`、`TABLE_GROUP_AMBIGUOUS`、`0x800A03EC`、`Number.Min` 未定義を検出し修正
- 失敗時の原因推定:
  - `Pdf.Tables` が表片を別ページ候補として返す
  - ヘッダー署名へデータ行や月トークンが混入する
  - Excel が文字列時刻を再解釈する

## 9. handoff 反映

本体変更後に handoff を再生成し、V2 の説明と配布物が本体の挙動とズレないことを確認しました。

## 10. 今後の改善候補

- `sameHeader` の署名ルールを profile から調整できるようにする
- `Review` の理由列を分類コード付きにする
- `24:00` 以外の日跨ぎ表記の扱いを定義する
- `現場名` / `氏名` の辞書正規化を次段階で追加する
- 負例 PDF をさらに追加して曖昧結合の境界を広げる
