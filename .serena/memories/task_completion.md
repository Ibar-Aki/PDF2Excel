# Task completion checklist
- Verify V1 is unaffected by any V2 change.
- Run relevant tests after changes. Minimum around launch/menu changes: unit tests plus BAT-related integration tests. For V2 table logic changes, run unit tests and targeted V2 integration tests at least.
- When reporting test results, include: 実施日時, 対象環境, 対象機能, 実行シナリオ, 結果概要, 所要時間, エラー有無, and failure cause estimate if failed.
- If template, profile, or handoff changes are made, keep distributed handoff artifacts in sync.
- Do not revert unrelated uncommitted changes.