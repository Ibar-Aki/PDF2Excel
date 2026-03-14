# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-14 23:04 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-14

## サマリー

- 実施日時: 2026-03-14 23:04:48 JST - 2026-03-14 23:04:49 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 8 件成功 / 0 件失敗
- 所要時間: 0 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString escapes quotes | PASS | 13 ms | A""B |
| 2 | ConvertTo-MTextListLiteral handles empty list | PASS | 7 ms | {} |
| 3 | ConvertTo-MLogicalLiteral returns lowercase literal | PASS | 6 ms | true / false |
| 4 | Get-ProfileOutputColumnNames uses prefix and source file | PASS | 11 ms | SourceFile,Column1,Column2,Column3 |
| 5 | Get-ProfileConfiguration loads default profile | PASS | 72 ms | 標準30列プロファイル / 30 |
| 6 | Get-StagingQueryFormula embeds folder path and staging source | PASS | 27 ms | Folder.Files + COLUMN_OVERFLOW confirmed |
| 7 | Get-ResultQueryFormula references staging query | PASS | 8 ms | Staging reference confirmed |
| 8 | Resolve-RunErrorInfo maps lock and cancel states | PASS | 4 ms | RUN_LOCKED / RUN_CANCELLED |
