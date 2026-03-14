# PDF2Excel 変更履歴

- 作成日: 2026-03-14 10:15 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-14

## 2026-03-14

### 追加

- 実行前チェック画面、帳票プロファイル機能、Summary 強化、エラー分類を追加
- 共通関数ファイル `scripts/pdf2excel.common.ps1` を追加
- VBA モジュールの Shift_JIS ミラー `template/vba/PDF2ExcelMacros.sjis.bas` を追加
- ユニットテスト `tests/run_unit_tests.ps1` を追加
- 日本語ファイル名テストと 50 件一括テストを追加
- 配布用最小パッケージに共通スクリプトを同梱
- `CHANGELOG.md` を追加

### 改善

- `KeepInput` を保管専用に見直し、実行ごとの staging 分離を導入
- 同時実行防止の mutex / lock file を導入
- 統合テストを子プロセス隔離とタイムアウト付きに改善
- Power Query の `Result` 側ロジックを `PDF2Excel_Staging` 参照へ統一
- `run_pdf2excel.ps1` を 1 ファイル維持のままセクション化し、内部 API を整理
- `System.Windows.Forms` の読み込みを遅延化し、headless 互換性を改善
- ログを 30 日超または 200 件超で自動整理するように改善
- テストレポートを日本語中心の表記へ統一
- README と主要ドキュメントのリンクを相対パス化し、GitHub / 配布先での可搬性を改善

### 修正

- 自己参照 `input` 実行時の削除事故を修正
- 同名 PDF の混在時に明示エラーで停止するよう修正
- VBA 取り込み方式を見直し、コンパイルエラー回避を強化
- Control / Summary シートの静的レイアウト定義の二重管理を解消
- BAT 末尾の余分な改行を除去
