# PDF2Excel セキュリティ監査レポート (現行実装レビュー)

- 作成日: 2026-03-20 16:35 JST
- 作成者: Codex (GPT-5)
- 評価対象: PDF2Excel プロジェクト現行実装 (`VER2 Secure` 導線中心)
- 評価基準日: 2026-03-20

注記:

- この文書は、現時点のコードベースに対する実装レビューです。
- 過去時点の厳しめ評価は `04-strict-security-audit-report.md` を参照してください。
- 現行運用の正解は `docs/01-current` 配下を優先してください。

## 1. エグゼクティブ・サマリ

現行の `VER2 Secure` 導線は、共有パス拒否、`LOCALAPPDATA` への runtime/logs 固定、ログのパスマスキング、`run.lock` の stale 判定、`AbandonedMutexException` からの自動回復など、直近の安定化改修によりかなり改善されています。  
現時点で、直ちに「運用停止が必要」と言える致命的脆弱性は確認していません。

一方で、コアスクリプトの信頼境界と監査証跡の堅牢性には、まだ詰めるべき点があります。特に、

- 外部指定プロファイルの受け入れ範囲が広いこと
- `run-history.csv` だけが非原子的に更新されていること

は、今後の保守運用を考えると早めに塞ぐべきです。

総評:

- 現状評価: `条件付きで実運用可能`
- 緊急度: `中`
- 優先対応件数: `2件`

## 2. 評価対象と確認方法

今回のレビューでは、主に次を確認しました。

- `scripts/run_pdf2excel.ps1`
- `scripts/run_pdf2excel_menu.ps1`
- `scripts/pdf2excel.common.ps1`
- `scripts/build_handoff_package.ps1`
- `README.md`
- `docs/01-current/*`

観点:

- 実行ポリシー
- runtime/logs の配置
- 多重起動防止
- 一時レポートと監査証跡の書き込み方式
- ログの情報露出
- 外部入力である profile 指定の信頼境界

## 3. 主要所見

### S-1: 外部指定プロファイルの信頼境界が広く、`ProfileName` も明示バリデーションがない

深刻度: Medium

事象:

- `RequestedProfilePath` は、存在する任意のパスであればそのまま `Resolve-Path` して読み込みます。
- `RequestedProfileName` は `\"$RequestedProfileName.json\"` へ直接連結され、許可文字やパストラバーサル相当の文字列に対する明示的な拒否がありません。

根拠:

- `scripts/run_pdf2excel.ps1:1199-1204`
- `scripts/run_pdf2excel.ps1:1206-1216`
- `scripts/run_pdf2excel.ps1:1219`

評価:

- メニューや通常運用では想定入力が比較的安全でも、コアスクリプト自体は CLI 直呼び出しを受け付けます。
- そのため、「任意の JSON を profile として食わせる」「profile 名にパス区切りや相対パスを混ぜる」余地をコア側で残しているのは、防御境界として弱いです。
- 直ちに任意コード実行へつながる形跡までは現時点で確認していませんが、Power Query 生成、列マッピング、出力位置制御の前提が profile へ強く依存しているため、想定外構成を安全に拒否するべきです。

推奨対応:

1. `ProfileName` は `^[A-Za-z0-9_-]+$` などの allowlist で制限する
2. `RequestedProfilePath` は既定では無効化し、必要時のみ明示スイッチで許可する
3. profile 読み込み前に「許可ディレクトリ配下か」を正規化パスで検証する
4. `profile-v2.schema.json` 相当の厳格バリデーションを追加する

### S-2: `run-history.csv` だけが非原子的に更新され、監査証跡が壊れる余地がある

深刻度: Medium

事象:

- `run-history.csv` は `Set-Content` / `Add-Content` で直接更新されています。
- `RunReport` と `environment-check.md` は原子的書き込みへ改善済みですが、履歴台帳だけ方式が揃っていません。

根拠:

- `scripts/run_pdf2excel.ps1:563-567`
- `scripts/run_pdf2excel.ps1:480-500`
- `scripts/run_pdf2excel.ps1:159`
- `scripts/run_pdf2excel.ps1:727`

評価:

- セキュリティ事故そのものというより、監査証跡の完全性の問題です。
- 電源断、AV 介入、複数プロセス競合、ファイルロックが重なった場合に CSV が途中破損すると、「いつ、誰が、何件処理したか」の証跡が崩れます。
- 運用監査やインシデント調査の観点では、これは軽視しない方がよいです。

推奨対応:

1. `run-history.csv` も temp file + replace の原子的書き込みへ統一する
2. 追記前に現行内容を読み直し、破損時は退避コピーを作ってから再生成する
3. 履歴整合チェックを環境チェックまたは保守コマンドへ追加する

## 4. 改善済みで評価できる点

次の点は、以前より明確に良くなっています。

### G-1: 共有パス実行を拒否し、runtime/logs をローカル固定している

- `VER2 Secure` 導線では、共有フォルダを作業領域として使わない方向へ寄っています。
- `%LOCALAPPDATA%\\PDF2Excel\\runtime` と `%LOCALAPPDATA%\\PDF2Excel\\logs` を使う設計は妥当です。

### G-2: ログのパスマスキングが実装されている

根拠:

- `scripts/pdf2excel.common.ps1:78`
- `scripts/pdf2excel.common.ps1:125`

評価:

- 既定ログで詳細パスを露出しない方針は、運用ログとして健全です。
- `DEBUG` 時だけ詳細を見る設計も妥当です。

### G-3: `RunReport` と環境チェックレポートは原子的書き込みへ改善済み

根拠:

- `scripts/run_pdf2excel.ps1:159`
- `scripts/run_pdf2excel.ps1:480-500`
- `scripts/run_pdf2excel.ps1:727`

評価:

- 以前の「途中書き込みを黙って読む」よりかなり堅くなっています。
- レポート系でこの方針が定着しているのは良いです。

### G-4: `AbandonedMutexException` 回復と stale `run.lock` 識別が入っている

根拠:

- `scripts/run_pdf2excel.ps1:586-604`
- `scripts/run_pdf2excel.ps1:626-705`
- `scripts/run_pdf2excel.ps1:988-1004`

評価:

- 前回異常終了後に次回実行まで巻き込んで失敗するリスクが減っています。
- 環境チェックで `実行中 / 前回異常終了の可能性 / 内容読取不可` を分けている点も保守運用上有効です。

### G-5: メニュー起動時の実行ポリシーは `RemoteSigned`

根拠:

- `scripts/run_pdf2excel_menu.ps1:94`
- `scripts/run_pdf2excel_menu.ps1:158`

評価:

- 少なくとも正式導線で `ExecutionPolicy Bypass` を常用していない点は評価できます。

## 5. 優先対応順

### 優先度 1

- profile 読み込み境界の厳格化
- `ProfileName` の allowlist 化
- `RequestedProfilePath` の既定無効化または保守者限定化

### 優先度 2

- `run-history.csv` の原子的更新
- 履歴破損時の自己診断

### 優先度 3

- セキュリティ観点では軽微ですが、`scripts/run_pdf2excel_menu.ps1` の手順書パスが旧 `docs\\user-manual.md` を向いているため、保守運用上は修正した方がよいです。

## 6. 結論

現行の `VER2 Secure` は、過去の厳しめ監査で問題だった領域のうち、共有実行、ログ露出、run lock 回復、レポート破損耐性といった運用系の弱点をかなり改善できています。  
一方で、コアスクリプトの入力信頼境界と履歴台帳の完全性には、まだ設計上の余白があります。

したがって、現時点の判断は次のとおりです。

- 今すぐ全面停止が必要な状態ではない
- ただし「入力境界の厳格化」と「監査証跡の原子化」は、次回の保守サイクルで優先実施すべき
- 次回監査では、profile schema 検証と履歴台帳の原子的更新が入っていることを合格条件にするのが妥当
