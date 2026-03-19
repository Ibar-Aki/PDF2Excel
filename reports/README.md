# PDF2Excel レポート運用ガイド

- 作成日: 2026-03-20 01:10 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## 目的

このファイルは、`reports/` 配下に何を置くか、何を Git 管理するかを明確にするための案内です。

## Git 管理するもの

- `test-report.md`
  - 最新の統合テスト結果レポートです。
- `unit-test-report.md`
  - 最新のユニットテスト結果レポートです。
- `.gitkeep`
  - 空ディレクトリ維持用です。

## Git 管理しない生成物

- `run-history.csv`
  - 実行履歴台帳です。利用端末ごとに内容が変わるため、Git 管理しません。
- `environment-check.md`
  - 環境チェック結果です。実行環境依存のため、Git 管理しません。

## 運用ルール

- コードや UI/UX を変えてテストを実施した場合は、`test-report.md` または `unit-test-report.md` を最新化します。
- `run-history.csv` と `environment-check.md` は調査用の一時成果物として扱い、コミット前に差分対象へ含めません。
- 追加のレポートを増やす場合は、このファイルと `docs/index.md` をあわせて更新します。
