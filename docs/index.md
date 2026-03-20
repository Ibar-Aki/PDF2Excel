# PDF2Excel ドキュメント一覧

- 作成日: 2026-03-13 00:51 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## はじめに

このページは、PDF2Excel の文書を目的別にたどれるようにした入口です。  
最初にどこを読めばよいか迷ったときは、このページから入ってください。

## 採番ルール

- `01-current` は現行運用で読むべき正本文書です。
- `02-guides` 以降は、補助ガイド、計画、報告、監査、内部メモを分けています。
- 各カテゴリ内の `01`, `02` は読む順番の目安です。

## まず見る最新版

- [README.md](../README.md)
  - 正式運用の入口、現在の機能、主要な注意点をまとめた概要です。
- [01-user-manual.md](01-current/01-user-manual.md)
  - 現在のメニュー操作と日常運用手順の最新版です。
- [03-project-layout.md](01-current/03-project-layout.md)
  - 現在のフォルダ構成と文書の置き場所を説明します。
- [04-technical-description.md](01-current/04-technical-description.md)
  - 現在の `VER2 Secure` 実装と主要な動作を説明します。
- [02-troubleshooting.md](01-current/02-troubleshooting.md)
  - 現在の障害切り分け手順です。

補足:

- `01-current` 配下を現行運用の正解として扱ってください。
- `03-plans`、`04-reports`、`05-security`、`90-internal` は作成時点の記録を含みます。

## 01-current

- [01-user-manual.md](01-current/01-user-manual.md)
  - 利用者向けの正式な手順書です。
- [02-troubleshooting.md](01-current/02-troubleshooting.md)
  - 症状別の確認方法と対処方法です。
- [03-project-layout.md](01-current/03-project-layout.md)
  - フォルダ構成と文書配置の説明です。
- [04-technical-description.md](01-current/04-technical-description.md)
  - 実装構造と処理フローの説明です。
- [05-maintenance-guide.md](01-current/05-maintenance-guide.md)
  - テンプレート更新、回帰確認、handoff 再生成を含む保守手順です。

## 02-guides

- [01-v2-sample-walkthrough.md](02-guides/01-v2-sample-walkthrough.md)
  - V2 サンプルを順番に試しながら `Result / Review / Errors` を体験できます。
- [02-v2-processing-flow-guide.md](02-guides/02-v2-processing-flow-guide.md)
  - V2 の処理フローを学習用に分解したガイドです。
- [03-excel-com-guide.md](02-guides/03-excel-com-guide.md)
  - Excel COM オートメーションの補助ガイドです。

## 03-plans

- [01-requirements-specification.md](03-plans/01-requirements-specification.md)
  - 現行システムの要件定義です。
- [02-improvement-proposals.md](03-plans/02-improvement-proposals.md)
  - 今後の改善提案を整理した文書です。
- [03-v2-data-transfer-proposal.md](03-plans/03-v2-data-transfer-proposal.md)
  - 生データ転記の考え方と PoC 方針です。
- [04-v2-remaining-issues-plan.md](03-plans/04-v2-remaining-issues-plan.md)
  - 当時の残課題計画です。
- [05-v3-new-repo-plan.md](03-plans/05-v3-new-repo-plan.md)
  - V3 を別レポジトリ化するための構成案と移行計画です。

## 04-reports

- [01-v2-tdd-implementation-report.md](04-reports/01-v2-tdd-implementation-report.md)
  - V2 の TDD 実装報告です。
- [02-v2-stabilization-implementation-report.md](04-reports/02-v2-stabilization-implementation-report.md)
  - V2 安定化改修の実装報告です。
- [03-review-remediation-report.md](04-reports/03-review-remediation-report.md)
  - レビュー指摘への対応記録です。
- [04-copilot-implementation-report.md](04-reports/04-copilot-implementation-report.md)
  - Copilot への実装委任手順の記録です。

## 05-security

- [01-security-risk-response-report.md](05-security/01-security-risk-response-report.md)
  - セキュリティリスク対応案の整理です。
- [02-system-risk-evaluation.md](05-security/02-system-risk-evaluation.md)
  - システム全体のリスク評価です。
- [03-security-audit-report.md](05-security/03-security-audit-report.md)
  - セキュリティ監査の詳細所見です。
- [04-strict-security-audit-report.md](05-security/04-strict-security-audit-report.md)
  - 厳しめ評価での監査所見です。
- [05-current-security-audit-report.md](05-security/05-current-security-audit-report.md)
  - 2026-03-20 時点の現行実装に対するセキュリティ監査です。
- [06-security-and-reliability-control-matrix.md](05-security/06-security-and-reliability-control-matrix.md)
  - セキュリティ、バグ、エラー対策を一覧表で整理した台帳です。

## 90-internal

- [01-codex-development-improvements.md](90-internal/01-codex-development-improvements.md)
  - Codex 利用時の改善メモと運用ノートです。

## 関連

- [CHANGELOG.md](../CHANGELOG.md)
  - 版ごとの主な変更点です。
- [samples/README.md](../samples/README.md)
  - `samples/common`、`samples/v2`、`legacy/v1/samples` の使い分けです。
- [reports/README.md](../reports/README.md)
  - `reports/` 配下の追跡対象と生成物の扱いです。

## おすすめの読み順

1. 全員共通で [README.md](../README.md)
2. 利用者は [01-user-manual.md](01-current/01-user-manual.md)
3. V2 を試すなら [01-v2-sample-walkthrough.md](02-guides/01-v2-sample-walkthrough.md)
4. 問題が起きたら [02-troubleshooting.md](01-current/02-troubleshooting.md)
5. 改修するなら [05-maintenance-guide.md](01-current/05-maintenance-guide.md)
