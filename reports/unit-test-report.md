# PDF2Excel Unit Test Report

- 作成日: 2026-03-14 08:57 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-14

## Summary

- 実施日時: 2026-03-14 08:57:16 JST - 2026-03-14 08:57:17 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 8 passed / 0 failed
- 所要時間: 1 秒
- エラー有無: なし

## Cases

| No | Test | Status | Duration | Notes |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString escapes quotes | PASS | 26 ms | A""B |
| 2 | ConvertTo-MTextListLiteral handles empty list | PASS | 23 ms | {} |
| 3 | ConvertTo-MLogicalLiteral returns lowercase literal | PASS | 38 ms | true / false |
| 4 | Get-ProfileOutputColumnNames uses prefix and source file | PASS | 27 ms | SourceFile,Column1,Column2,Column3 |
| 5 | Get-ProfileConfiguration loads default profile | PASS | 152 ms | 標準30列プロファイル / 30 |
| 6 | Get-StagingQueryFormula embeds folder path and staging source | PASS | 81 ms | Folder.Files + COLUMN_OVERFLOW confirmed |
| 7 | Get-ResultQueryFormula references staging query | PASS | 26 ms | Staging reference confirmed |
| 8 | Resolve-RunErrorInfo maps lock and cancel states | PASS | 24 ms | RUN_LOCKED / RUN_CANCELLED |
