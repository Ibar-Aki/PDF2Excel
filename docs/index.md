# PDF2Excel ドキュメント一覧

- 作成日: 2026-03-13 00:51 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## はじめに

このページは、PDF2Excel の文書を目的別にたどれるようにした案内ページです。  
最初にどこを読めばよいか迷ったときは、このページから入ってください。

## まず見る最新版

- [README.md](../README.md)
  - 正式運用の入口、現在の機能、主要な注意点をまとめた最新版です。
- [user-manual.md](user-manual.md)
  - 現在のメニュー操作、`[0] 環境チェック`、`[6] VER2 設定ウィザード` を含む利用手順の最新版です。
- [project-layout.md](project-layout.md)
  - フォルダ構成、生成物の置き場所、`reports/` の扱いを含む構成案内の最新版です。
- [technical-description.md](technical-description.md)
  - 現在の `VER2 Secure` 構成、環境チェック、実行履歴、テンプレート整合性チェックを含む技術説明の最新版です。
- [troubleshooting.md](troubleshooting.md)
  - 現在の障害切り分け手順です。環境チェックと実行履歴の見方も含みます。

補足:

- 監査レポート、実装報告、計画書は「作成日時点の記録」です。
- 現行運用の正解は、原則として上の最新版 5 文書を優先してください。

## 利用者向け

- [README.md](../README.md)
  - 最短で全体像をつかむための概要です。
- [CHANGELOG.md](../CHANGELOG.md)
  - 版ごとの主な変更点を確認できます。
- [user-manual.md](user-manual.md)
  - 初回利用、日常運用、確認ポイントまで含めた詳しい手順書です。
- [project-layout.md](project-layout.md)
  - フォルダ構成と、どのファイルが何のためにあるかを説明しています。
- [improvement-proposals.md](improvement-proposals.md)
  - 今後の改善余地を、優先度と効果つきで整理した提案書です。
- [copilot-implementation-report.md](copilot-implementation-report.md)
  - Microsoft Copilot に段階実装させるための進め方とプロンプト雛形です。
- [requirements-specification.md](requirements-specification.md)
  - 現行システムの要件、対象範囲、機能要件、非機能要件を整理した要件定義書です。
- [technical-description.md](technical-description.md)
  - 動作フロー、技術スタック、アーキテクチャ、主要モジュールを整理した技術説明書です。
- [v2-data-transfer-proposal.md](v2-data-transfer-proposal.md)
  - 生データ転記方式と PoC 方針を整理した提案書です。
- [v2-tdd-implementation-report.md](v2-tdd-implementation-report.md)
  - V2 の多ページ結合、時刻正規化、TDD 実施内容、代表検証結果をまとめた実装報告です。
- [v2-stabilization-implementation-report.md](v2-stabilization-implementation-report.md)
  - V2 安定化改修で行った調査、設計判断、実装順序、テスト結果、handoff 反映までをまとめた詳細報告です。
- [review-remediation-report.md](review-remediation-report.md)
  - レビュー指摘に対して何を直したか、不要判断の有無も含めて整理した記録です。
- [codex-development-improvements.md](codex-development-improvements.md)
  - Codex 利用時の改善点、AGENTS ルール、公式 skill の扱いを整理したメモです。

## 障害対応向け

- [reports/README.md](../reports/README.md)
  - `reports/` 配下の追跡対象と生成物の扱いを説明します。
- [troubleshooting.md](troubleshooting.md)
  - 症状別の確認方法と対処方法をまとめています。
- [test-report.md](../reports/test-report.md)
  - 最新の統合テスト結果です。
- [unit-test-report.md](../reports/unit-test-report.md)
  - 共通関数とクエリ生成のユニットテスト結果です。

## 監査・評価

- [security-risk-response-report.md](security-risk-response-report.md)
  - セキュリティ上の論点と、採用済み対策の整理です。
- [system-risk-evaluation.md](system-risk-evaluation.md)
  - システム全体のリスク評価です。
- [security-audit-report.md](security-audit-report.md)
  - セキュリティ監査の詳細所見です。
- [strict-security-audit-report.md](strict-security-audit-report.md)
  - 厳しめ評価での監査所見です。

## 保守・改修向け

- [maintenance-guide.md](maintenance-guide.md)
  - テンプレート更新、テスト、リリース前確認のための保守手順です。
- [requirements-specification.md](requirements-specification.md)
  - 仕様確認や変更影響の整理を行うときの基準文書です。
- [technical-description.md](technical-description.md)
  - 実装構造と処理のつながりを把握するための技術文書です。
- [run_pdf2excel.ps1](../scripts/run_pdf2excel.ps1)
  - 変換処理の本体です。
- [build_excel_template.ps1](../scripts/build_excel_template.ps1)
  - Excel テンプレートを再生成します。
- [PDF2ExcelMacros.bas](../template/vba/PDF2ExcelMacros.bas)
  - VBA マクロ本体です。
- [PDF2ExcelMacros.sjis.bas](../template/vba/PDF2ExcelMacros.sjis.bas)
  - VBA モジュールの Shift_JIS ミラーです。UTF-8 版を正本として同期します。
- [run_integration_tests.ps1](../tests/run_integration_tests.ps1)
  - 統合テストの実行スクリプトです。
- [run_unit_tests.ps1](../tests/run_unit_tests.ps1)
  - 主要な共通関数とクエリ生成のユニットテストを実行します。

## おすすめの読み順

1. 全員共通で [README.md](../README.md)
2. 利用者は [user-manual.md](user-manual.md)
3. 問題が起きたら [troubleshooting.md](troubleshooting.md)
4. 改修するなら [maintenance-guide.md](maintenance-guide.md)
