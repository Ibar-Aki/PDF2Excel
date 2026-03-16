# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-17 00:05 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-17

## サマリー

- 実施日時: 2026-03-17 00:05:58 JST - 2026-03-17 00:05:59 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 40 件成功 / 0 件失敗
- 所要時間: 1 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString は二重引用符をエスケープする | 成功 | 22 ms | A""B |
| 2 | ConvertTo-MTextListLiteral は空配列を処理できる | 成功 | 30 ms | {} |
| 3 | ConvertTo-MLogicalLiteral は小文字の真偽値を返す | 成功 | 6 ms | true / false |
| 4 | Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる | 成功 | 19 ms | SourceFile,Column1,Column2,Column3 |
| 5 | Get-ProfileConfiguration は既定プロファイルを読み込む | 成功 | 119 ms | 標準30列プロファイル / 30 |
| 6 | Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む | 成功 | 44 ms | Folder.Files と COLUMN_OVERFLOW を確認 |
| 7 | Get-ResultQueryFormula は staging クエリを参照する | 成功 | 10 ms | Staging 参照を確認 |
| 8 | Resolve-RunErrorInfo はロックとキャンセルを分類する | 成功 | 6 ms | RUN_LOCKED / RUN_CANCELLED |
| 9 | Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む | 成功 | 5 ms | 日本語勤怠管理表プロファイル / 35 |
| 10 | Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む | 成功 | 5 ms | 日本語売上日報プロファイル / 12 |
| 11 | Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む | 成功 | 4 ms | 日本語在庫一覧プロファイル / 10 |
| 12 | Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む | 成功 | 9 ms | 日本語問い合わせ管理表プロファイル / 9 |
| 13 | Get-ProfileConfiguration は生データ転記PoCプロファイルを読み込む | 成功 | 13 ms | 生データ転記PoCプロファイル / 30 |
| 14 | TemplateBuilder VBA は V2 テンプレート生成マクロを持つ | 成功 | 4 ms | Review / BuildPDF2ExcelV2 / 5 シート / 重複除去を確認 |
| 15 | テンプレート再作成手順は配置先を明記する | 成功 | 28 ms | README / user-manual / handoff README の配置先明記を確認 |
| 16 | run_pdf2excel.ps1 は Secure モードでローカル runtime を使う | 成功 | 3 ms | SecurityMode / LOCALAPPDATA / KeepInput 無効化を確認 |
| 17 | V2 ラッパーは Secure モードを渡し RemoteSigned で起動する | 成功 | 5 ms | SecurityMode Secure / RemoteSigned / Bypass 除去を確認 |
| 18 | V2 BAT は ASCII かつ RemoteSigned で起動する | 成功 | 9 ms | ASCII / RemoteSigned / Bypass 除去を確認 |
| 19 | V2 BAT とメニューは ForceMenu 導線を持つ | 成功 | 5 ms | ForceMenu 導線を確認 |
| 20 | 共通メニューは完了メッセージを版別に分ける | 成功 | 20 ms | V1/V2 完了メッセージ分岐を確認 |
| 21 | handoff ビルドは VBA モジュールを同梱する | 成功 | 19 ms | template\vba / sjis ミラー同期 / TemplateBuilder.bas / PDF2ExcelMacros.bas を確認 |
| 22 | build_handoff_package は V2 限定再生成を受け付ける | 成功 | 5 ms | TargetVersion フィルタを確認 |
| 23 | run_pdf2excel.ps1 は待機メッセージを表示する | 成功 | 56 ms | 待機メッセージを確認 |
| 24 | Normalize-TimeText は全角コロンを半角時刻へ正規化する | 成功 | 43 ms | 08:00 / 480 |
| 25 | Normalize-TimeText は時分表記と空白込みを正規化する | 成功 | 1 ms | 09:15 / 555 |
| 26 | Normalize-TimeText は時のみ表記を 00 分補完する | 成功 | 1 ms | 18:00 / 1080 |
| 27 | Normalize-TimeText は Excel 時刻比率の文字列も正規化する | 成功 | 28 ms | 09:15 / 555 |
| 28 | Normalize-TimeText は 24:00 を有効時刻として扱う | 成功 | 2 ms | 24:00 / 1440 |
| 29 | Normalize-TimeText は 24:30 を範囲外扱いにする | 成功 | 2 ms | INVALID / 時刻の範囲外です。 |
| 30 | Normalize-TimeText は 整数文字列 1 を 01:00 として扱う | 成功 | 1 ms | 01:00 / 60 |
| 31 | Normalize-TimeText は 24:00:00 を 24:00 として扱う | 成功 | 8 ms | 24:00 / 1440 |
| 32 | Get-ResultOutputColumnNames は V2 で正規化列を追加する | 成功 | 23 ms | 正規化入場2_分,正規化退場2,正規化退場2_分,時刻正規化状態,時刻確認メモ |
| 33 | Get-NormalizedTimeColumnDefinitions は Review raw 列名を返す | 成功 | 52 ms | 正規化入場1_raw,正規化退場1_raw |
| 34 | Get-ReviewReasonCategories は HEADER_MISMATCH を先頭に重複なく返す | 成功 | 68 ms | HEADER_MISMATCH,TIME_MULTI,TIME_MISSING,TIME_INVALID |
| 35 | Get-StagingQueryFormula は V2 で canonical sameHeader と曖昧分離を考慮する | 成功 | 49 ms | canonical sameHeader / ambiguity split / review raw / reason categories |
| 36 | Get-ReviewQueryFormulaV2 は ReasonCategory を展開対象に含める | 成功 | 13 ms | ReasonCategory expansion ready |
| 37 | run_pdf2excel.bat は ASCII のみで構成される | 成功 | 6 ms | ASCII のみを確認 |
| 38 | run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される | 成功 | 14 ms | 版別 BAT の ASCII を確認 |
| 39 | run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される | 成功 | 7 ms | UTF-8 BOM を確認 |
| 40 | 版別メニュー PowerShell は UTF-8 BOM で保存される | 成功 | 3 ms | 版別メニューの UTF-8 BOM を確認 |
