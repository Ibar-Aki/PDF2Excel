# AGENTS.md

- 作成日: 2026-03-15 11:05 JST
- 作成者: Codex (GPT-5)

## プロジェクトの目的

このディレクトリには、PowerShell / Excel(M365) / Power Query を使ってテキスト PDF の表を Excel へ転記するローカルツールが含まれています。

## 作業ルール

- 利用者向けの BAT は ASCII のみで構成してください。
- 日本語の表示文言は BAT に直書きせず、UTF-8 BOM の PowerShell スクリプト側へ置いてください。
- 日本語を含む新規 `.ps1` を作る場合は、PowerShell 5.1 互換のため UTF-8 BOM を前提にしてください。
- `run_pdf2excel.bat` を変更した場合は、`handoff/generated/PDF2Excel_V2_Minimal/run_pdf2excel.bat` と `scripts/build_handoff_package.ps1` への影響を確認してください。
- `scripts/run_pdf2excel_menu.ps1` を変更した場合は、BAT 直実行の回帰確認を行ってください。
- 新しい帳票プロファイルを追加した場合は、対応するサンプル PDF と少なくとも 1 件のテストを追加してください。
- `samples/common/pdf` 配下の既存サンプルを意図せず削除しないでください。生成スクリプトは管理対象ディレクトリだけを更新してください。
- UI/UX に関わる変更を行った場合は、`README.md` または `docs/user-manual.md` に必ず記録してください。
- テスト結果を報告する場合は、実施日時、対象環境、対象機能、シナリオ、結果概要、所要時間、エラー有無を含めてください。

## 推奨確認

- BAT / 起動まわり変更時
  - `tests/run_unit_tests.ps1`
  - `tests/run_integration_tests.ps1 -CaseName 'BAT 経由の変換'`
  - `tests/run_integration_tests.ps1 -CaseName 'BAT 直実行で待機しない'`
- サンプル生成変更時
  - `scripts/build_sample_pdfs.ps1`
  - サンプルの生成後に既存サンプルが残っていることを確認
- 新規プロファイル追加時
  - `tests/run_unit_tests.ps1`
  - 対応する統合テスト 1 件以上
