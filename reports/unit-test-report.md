# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-19 07:52 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-19

## サマリー

- 実施日時: 2026-03-19 07:52:31 JST - 2026-03-19 07:52:32 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 49 件成功 / 0 件失敗
- 所要時間: 1 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString は二重引用符をエスケープする | 成功 | 10 ms | A""B |
| 2 | ConvertTo-MTextListLiteral は空配列を処理できる | 成功 | 7 ms | {} |
| 3 | ConvertTo-MLogicalLiteral は小文字の真偽値を返す | 成功 | 4 ms | true / false |
| 4 | Get-PathLocationInfo は UNC とローカルパスを判定できる | 成功 | 6 ms | UNC / local を確認 |
| 5 | Get-LocalAppDataPdf2ExcelPath は PDF2Excel 配下を返す | 成功 | 3 ms | C:\Users\TestUser\AppData\Local\PDF2Excel\logs |
| 6 | Protect-MessagePaths は登録済みパスをマスクできる | 成功 | 27 ms | 出力先=...\logs\run_001.log / staging=...\staging\a.pdf |
| 7 | Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる | 成功 | 9 ms | SourceFile,Column1,Column2,Column3 |
| 8 | Get-ProfileConfiguration は既定プロファイルを読み込む | 成功 | 75 ms | 標準30列プロファイル / 30 |
| 9 | Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む | 成功 | 28 ms | Folder.Files と COLUMN_OVERFLOW を確認 |
| 10 | Get-ResultQueryFormula は staging クエリを参照する | 成功 | 9 ms | Staging 参照を確認 |
| 11 | Resolve-RunErrorInfo はロックとキャンセルを分類する | 成功 | 5 ms | RUN_LOCKED / RUN_CANCELLED |
| 12 | Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む | 成功 | 4 ms | 日本語勤怠管理表プロファイル / 35 |
| 13 | Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む | 成功 | 4 ms | 日本語売上日報プロファイル / 12 |
| 14 | Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む | 成功 | 4 ms | 日本語在庫一覧プロファイル / 10 |
| 15 | Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む | 成功 | 4 ms | 日本語問い合わせ管理表プロファイル / 9 |
| 16 | Get-ProfileConfiguration は生データ転記サンプルプロファイルを読み込む | 成功 | 10 ms | 生データ転記サンプルプロファイル / 30 |
| 17 | TemplateBuilder VBA は V2 テンプレート生成マクロを持つ | 成功 | 2 ms | Review / BuildPDF2ExcelV2 / 5 シート / 重複除去を確認 |
| 18 | テンプレート再作成手順は配置先を明記する | 成功 | 47 ms | README / user-manual / handoff README の配置先明記を確認 |
| 19 | run_pdf2excel.ps1 は Secure モードでローカル runtime を使う | 成功 | 2 ms | SecurityMode / LogLevel / LOCALAPPDATA runtime/logs / path mask / force stop / KeepInput 無効化を確認 |
| 20 | V2 ラッパーは Secure モードを渡し RemoteSigned で起動する | 成功 | 2 ms | SecurityMode Secure / RemoteSigned / Bypass 除去を確認 |
| 21 | V2 BAT は ASCII かつ RemoteSigned で起動する | 成功 | 6 ms | ASCII / RemoteSigned / Bypass 除去を確認 |
| 22 | 正式運用 BAT は V2 Secure を起動し RemoteSigned に統一されている | 成功 | 4 ms | run_pdf2excel.bat は V2 Secure、開発導線も RemoteSigned を確認 |
| 23 | テストとビルド導線に Bypass が残っていない | 成功 | 20 ms | tests / build 呼び出しの Bypass 除去を確認 |
| 24 | V2 BAT とメニューは ForceMenu 導線を持つ | 成功 | 3 ms | ForceMenu 導線を確認 |
| 25 | 共通メニューは共有パス制約とプロファイル雛形導線を持つ | 成功 | 1 ms | 共有パス制約 / LOCALAPPDATA 必須 / プロファイル雛形 / メニュー番号を確認 |
| 26 | 共通メニューは完了メッセージを版別に分ける | 成功 | 1 ms | V1/V2 完了メッセージ分岐を確認 |
| 27 | handoff ビルドは VBA モジュールを同梱する | 成功 | 1 ms | template\vba / sjis ミラー同期 / new_profile_scaffold.ps1 / TemplateBuilder.bas / PDF2ExcelMacros.bas を確認 |
| 28 | new_profile_scaffold は v1/v2 雛形を生成できる | 成功 | 750 ms |  VER1 のプロファイル雛形を作成しました。 内部ID: sample_v1 表示名: テストV1 保存先: C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v1.json  VER2 のプロファイル雛形を作成しました。 内部ID: sample_v2 表示名: テストV2 保存先: C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v2.json v1/v2 雛形生成を確認 |
| 29 | build_handoff_package は V2 限定再生成を受け付ける | 成功 | 1 ms | TargetVersion フィルタを確認 |
| 30 | run.lock は匿名化され、Secure cleanup 失敗時の警告を持つ | 成功 | 2 ms | run.lock 匿名化 / cleanup 警告強化を確認 |
| 31 | 日本語コンソール出力スクリプトは UTF-8 初期化と保存形式を持つ | 成功 | 21 ms | run/build/scaffold scripts の UTF-8 初期化と build_handoff_package.ps1 の BOM を確認 |
| 32 | run_pdf2excel.ps1 は待機メッセージを表示する | 成功 | 2 ms | 待機メッセージを確認 |
| 33 | Normalize-TimeText は全角コロンを半角時刻へ正規化する | 成功 | 30 ms | 08:00 / 480 |
| 34 | Normalize-TimeText は時分表記と空白込みを正規化する | 成功 | 1 ms | 09:15 / 555 |
| 35 | Normalize-TimeText は時のみ表記を 00 分補完する | 成功 | 1 ms | 18:00 / 1080 |
| 36 | Normalize-TimeText は Excel 時刻比率の文字列も正規化する | 成功 | 20 ms | 09:15 / 555 |
| 37 | Normalize-TimeText は 24:00 を有効時刻として扱う | 成功 | 1 ms | 24:00 / 1440 |
| 38 | Normalize-TimeText は 24:30 を範囲外扱いにする | 成功 | 2 ms | INVALID / 時刻の範囲外です。 |
| 39 | Normalize-TimeText は 整数文字列 1 を 01:00 として扱う | 成功 | 1 ms | 01:00 / 60 |
| 40 | Normalize-TimeText は 24:00:00 を 24:00 として扱う | 成功 | 5 ms | 24:00 / 1440 |
| 41 | Get-ResultOutputColumnNames は V2 で正規化列を追加する | 成功 | 17 ms | 正規化入場2_分,正規化退場2,正規化退場2_分,時刻正規化状態,時刻確認メモ |
| 42 | Get-NormalizedTimeColumnDefinitions は Review raw 列名を返す | 成功 | 11 ms | 正規化入場1_raw,正規化退場1_raw |
| 43 | Get-ReviewReasonCategories は HEADER_MISMATCH を先頭に重複なく返す | 成功 | 51 ms | HEADER_MISMATCH,TIME_MULTI,TIME_MISSING,TIME_INVALID |
| 44 | Get-StagingQueryFormula は V2 で canonical sameHeader と曖昧分離を考慮する | 成功 | 31 ms | canonical sameHeader / ambiguity split / review raw / reason categories |
| 45 | Get-ReviewQueryFormulaV2 は ReasonCategory を展開対象に含める | 成功 | 9 ms | ReasonCategory expansion ready |
| 46 | run_pdf2excel.bat は ASCII のみで構成される | 成功 | 8 ms | ASCII のみを確認 |
| 47 | run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される | 成功 | 11 ms | 版別 BAT の ASCII を確認 |
| 48 | run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される | 成功 | 1 ms | UTF-8 BOM を確認 |
| 49 | 版別メニュー PowerShell は UTF-8 BOM で保存される | 成功 | 2 ms | 版別メニューの UTF-8 BOM を確認 |
