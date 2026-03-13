# PDF2Excel ドキュメント一覧

- 作成日: 2026-03-13 00:51 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-14

## はじめに

このページは、PDF2Excel の文書を目的別にたどれるようにした案内ページです。  
最初にどこを読めばよいか迷ったときは、このページから入ってください。

## 利用者向け

- [README.md](C:\Work_Codex\PDF2Excel\README.md)
  - 最短で全体像をつかむための概要です。
- [user-manual.md](C:\Work_Codex\PDF2Excel\docs\user-manual.md)
  - 初回利用、日常運用、確認ポイントまで含めた詳しい手順書です。
- [project-layout.md](C:\Work_Codex\PDF2Excel\docs\project-layout.md)
  - フォルダ構成と、どのファイルが何のためにあるかを説明しています。
- [improvement-proposals.md](C:\Work_Codex\PDF2Excel\docs\improvement-proposals.md)
  - 今後の改善余地を、優先度と効果つきで整理した提案書です。
- [copilot-implementation-report.md](C:\Work_Codex\PDF2Excel\docs\copilot-implementation-report.md)
  - Microsoft Copilot に段階実装させるための進め方とプロンプト雛形です。
- [requirements-specification.md](C:\Work_Codex\PDF2Excel\docs\requirements-specification.md)
  - 現行システムの要件、対象範囲、機能要件、非機能要件を整理した要件定義書です。
- [technical-description.md](C:\Work_Codex\PDF2Excel\docs\technical-description.md)
  - 動作フロー、技術スタック、アーキテクチャ、主要モジュールを整理した技術説明書です。

## 障害対応向け

- [troubleshooting.md](C:\Work_Codex\PDF2Excel\docs\troubleshooting.md)
  - 症状別の確認方法と対処方法をまとめています。
- [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md)
  - 最新の統合テスト結果です。

## 保守・改修向け

- [maintenance-guide.md](C:\Work_Codex\PDF2Excel\docs\maintenance-guide.md)
  - テンプレート更新、テスト、リリース前確認のための保守手順です。
- [requirements-specification.md](C:\Work_Codex\PDF2Excel\docs\requirements-specification.md)
  - 仕様確認や変更影響の整理を行うときの基準文書です。
- [technical-description.md](C:\Work_Codex\PDF2Excel\docs\technical-description.md)
  - 実装構造と処理のつながりを把握するための技術文書です。
- [run_pdf2excel.ps1](C:\Work_Codex\PDF2Excel\scripts\run_pdf2excel.ps1)
  - 変換処理の本体です。
- [build_excel_template.ps1](C:\Work_Codex\PDF2Excel\scripts\build_excel_template.ps1)
  - Excel テンプレートを再生成します。
- [PDF2ExcelMacros.bas](C:\Work_Codex\PDF2Excel\template\vba\PDF2ExcelMacros.bas)
  - VBA マクロ本体です。
- [run_integration_tests.ps1](C:\Work_Codex\PDF2Excel\tests\run_integration_tests.ps1)
  - 統合テストの実行スクリプトです。

## おすすめの読み順

1. 全員共通で [README.md](C:\Work_Codex\PDF2Excel\README.md)
2. 利用者は [user-manual.md](C:\Work_Codex\PDF2Excel\docs\user-manual.md)
3. 問題が起きたら [troubleshooting.md](C:\Work_Codex\PDF2Excel\docs\troubleshooting.md)
4. 改修するなら [maintenance-guide.md](C:\Work_Codex\PDF2Excel\docs\maintenance-guide.md)
