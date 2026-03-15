# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-15 15:19 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-15

## サマリー

- 実施日時: 2026-03-15 15:19:32 JST - 2026-03-15 15:19:33 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 23 件成功 / 0 件失敗
- 所要時間: 1 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString は二重引用符をエスケープする | 成功 | 18 ms | A""B |
| 2 | ConvertTo-MTextListLiteral は空配列を処理できる | 成功 | 30 ms | {} |
| 3 | ConvertTo-MLogicalLiteral は小文字の真偽値を返す | 成功 | 24 ms | true / false |
| 4 | Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる | 成功 | 32 ms | SourceFile,Column1,Column2,Column3 |
| 5 | Get-ProfileConfiguration は既定プロファイルを読み込む | 成功 | 153 ms | 標準30列プロファイル / 30 |
| 6 | Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む | 成功 | 118 ms | Folder.Files と COLUMN_OVERFLOW を確認 |
| 7 | Get-ResultQueryFormula は staging クエリを参照する | 成功 | 29 ms | Staging 参照を確認 |
| 8 | Resolve-RunErrorInfo はロックとキャンセルを分類する | 成功 | 14 ms | RUN_LOCKED / RUN_CANCELLED |
| 9 | Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む | 成功 | 5 ms | 日本語勤怠管理表プロファイル / 35 |
| 10 | Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む | 成功 | 5 ms | 日本語売上日報プロファイル / 12 |
| 11 | Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む | 成功 | 21 ms | 日本語在庫一覧プロファイル / 10 |
| 12 | Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む | 成功 | 6 ms | 日本語問い合わせ管理表プロファイル / 9 |
| 13 | Get-ProfileConfiguration は建設現場転記PoCプロファイルを読み込む | 成功 | 35 ms | 建設現場転記PoCプロファイル / 30 |
| 14 | Normalize-TimeText は全角コロンを半角時刻へ正規化する | 成功 | 71 ms | 08:00 / 480 |
| 15 | Normalize-TimeText は時分表記と空白込みを正規化する | 成功 | 20 ms | 09:15 / 555 |
| 16 | Normalize-TimeText は時のみ表記を 00 分補完する | 成功 | 6 ms | 18:00 / 1080 |
| 17 | Normalize-TimeText は Excel 時刻比率の文字列も正規化する | 成功 | 21 ms | 09:15 / 555 |
| 18 | Get-ResultOutputColumnNames は V2 で正規化列を追加する | 成功 | 76 ms | 正規化入場2_分,正規化退場2,正規化退場2_分,時刻正規化状態,時刻確認メモ |
| 19 | Get-StagingQueryFormula は V2 で sameHeader とページ番号を考慮する | 成功 | 48 ms | sameHeader / __PageNumber / review note |
| 20 | run_pdf2excel.bat は ASCII のみで構成される | 成功 | 41 ms | ASCII のみを確認 |
| 21 | run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される | 成功 | 9 ms | 版別 BAT の ASCII を確認 |
| 22 | run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される | 成功 | 20 ms | UTF-8 BOM を確認 |
| 23 | 版別メニュー PowerShell は UTF-8 BOM で保存される | 成功 | 19 ms | 版別メニューの UTF-8 BOM を確認 |
