# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-16 00:08 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-16

## サマリー

- 実施日時: 2026-03-16 00:08:57 JST - 2026-03-16 00:08:58 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 30 件成功 / 0 件失敗
- 所要時間: 1 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString は二重引用符をエスケープする | 成功 | 22 ms | A""B |
| 2 | ConvertTo-MTextListLiteral は空配列を処理できる | 成功 | 10 ms | {} |
| 3 | ConvertTo-MLogicalLiteral は小文字の真偽値を返す | 成功 | 8 ms | true / false |
| 4 | Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる | 成功 | 22 ms | SourceFile,Column1,Column2,Column3 |
| 5 | Get-ProfileConfiguration は既定プロファイルを読み込む | 成功 | 143 ms | 標準30列プロファイル / 30 |
| 6 | Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む | 成功 | 37 ms | Folder.Files と COLUMN_OVERFLOW を確認 |
| 7 | Get-ResultQueryFormula は staging クエリを参照する | 成功 | 11 ms | Staging 参照を確認 |
| 8 | Resolve-RunErrorInfo はロックとキャンセルを分類する | 成功 | 6 ms | RUN_LOCKED / RUN_CANCELLED |
| 9 | Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む | 成功 | 6 ms | 日本語勤怠管理表プロファイル / 35 |
| 10 | Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む | 成功 | 5 ms | 日本語売上日報プロファイル / 12 |
| 11 | Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む | 成功 | 5 ms | 日本語在庫一覧プロファイル / 10 |
| 12 | Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む | 成功 | 5 ms | 日本語問い合わせ管理表プロファイル / 9 |
| 13 | Get-ProfileConfiguration は建設現場転記PoCプロファイルを読み込む | 成功 | 13 ms | 建設現場転記PoCプロファイル / 30 |
| 14 | Normalize-TimeText は全角コロンを半角時刻へ正規化する | 成功 | 38 ms | 08:00 / 480 |
| 15 | Normalize-TimeText は時分表記と空白込みを正規化する | 成功 | 3 ms | 09:15 / 555 |
| 16 | Normalize-TimeText は時のみ表記を 00 分補完する | 成功 | 1 ms | 18:00 / 1080 |
| 17 | Normalize-TimeText は Excel 時刻比率の文字列も正規化する | 成功 | 30 ms | 09:15 / 555 |
| 18 | Normalize-TimeText は 24:00 を有効時刻として扱う | 成功 | 2 ms | 24:00 / 1440 |
| 19 | Normalize-TimeText は 24:30 を範囲外扱いにする | 成功 | 11 ms | INVALID / 時刻の範囲外です。 |
| 20 | Normalize-TimeText は 整数文字列 1 を 01:00 として扱う | 成功 | 2 ms | 01:00 / 60 |
| 21 | Normalize-TimeText は 24:00:00 を 24:00 として扱う | 成功 | 14 ms | 24:00 / 1440 |
| 22 | Get-ResultOutputColumnNames は V2 で正規化列を追加する | 成功 | 91 ms | 正規化入場2_分,正規化退場2,正規化退場2_分,時刻正規化状態,時刻確認メモ |
| 23 | Get-NormalizedTimeColumnDefinitions は Review raw 列名を返す | 成功 | 22 ms | 正規化入場1_raw,正規化退場1_raw |
| 24 | Get-ReviewReasonCategories は HEADER_MISMATCH を先頭に重複なく返す | 成功 | 91 ms | HEADER_MISMATCH,TIME_MULTI,TIME_MISSING,TIME_INVALID |
| 25 | Get-StagingQueryFormula は V2 で canonical sameHeader と曖昧分離を考慮する | 成功 | 75 ms | canonical sameHeader / ambiguity split / review raw / reason categories |
| 26 | Get-ReviewQueryFormulaV2 は ReasonCategory を展開対象に含める | 成功 | 11 ms | ReasonCategory expansion ready |
| 27 | run_pdf2excel.bat は ASCII のみで構成される | 成功 | 8 ms | ASCII のみを確認 |
| 28 | run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される | 成功 | 10 ms | 版別 BAT の ASCII を確認 |
| 29 | run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される | 成功 | 6 ms | UTF-8 BOM を確認 |
| 30 | 版別メニュー PowerShell は UTF-8 BOM で保存される | 成功 | 3 ms | 版別メニューの UTF-8 BOM を確認 |
