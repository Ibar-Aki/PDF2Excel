# PDF2Excel ユニットテストレポート

- 作成日: 2026-03-20 17:37 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

## サマリー

- 実施日時: 2026-03-20 17:37:59 JST - 2026-03-20 17:38:02 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920
- 対象機能: 共通関数、プロファイル解決、Power Query 文字列生成
- 結果概要: 68 件成功 / 0 件失敗
- 所要時間: 3 秒
- エラー有無: なし

## テスト結果

| No | テスト | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Escape-MString は二重引用符をエスケープする | 成功 | 13 ms | A""B |
| 2 | ConvertTo-MTextListLiteral は空配列を処理できる | 成功 | 8 ms | {} |
| 3 | ConvertTo-MLogicalLiteral は小文字の真偽値を返す | 成功 | 7 ms | true / false |
| 4 | Get-PathLocationInfo は UNC とローカルパスを判定できる | 成功 | 8 ms | UNC / local を確認 |
| 5 | Get-LocalAppDataPdf2ExcelPath は PDF2Excel 配下を返す | 成功 | 5 ms | C:\Users\TestUser\AppData\Local\PDF2Excel\logs |
| 6 | Protect-MessagePaths は登録済みパスをマスクできる | 成功 | 54 ms | 出力先=...\logs\run_001.log / staging=...\staging\a.pdf |
| 7 | Get-ProfileOutputColumnNames は接頭辞と元ファイル列を並べる | 成功 | 16 ms | SourceFile,Column1,Column2,Column3 |
| 8 | Get-ProfileConfiguration は既定プロファイルを読み込む | 成功 | 152 ms | 標準30列プロファイル / 30 |
| 9 | Get-StagingQueryFormula は入力フォルダと判定ロジックを埋め込む | 成功 | 35 ms | Folder.Files と COLUMN_OVERFLOW を確認 |
| 10 | Get-ResultQueryFormula は staging クエリを参照する | 成功 | 75 ms | Staging 参照を確認 |
| 11 | Resolve-RunErrorInfo はロックとキャンセルを分類する | 成功 | 7 ms | RUN_LOCKED / RUN_CANCELLED |
| 12 | Resolve-RunErrorInfo はプロファイルと LOCALAPPDATA を分類する | 成功 | 4 ms | PROFILE_RESOLUTION_ERROR / LOCAL_RUNTIME_UNAVAILABLE |
| 13 | Get-RunErrorGuidance は次アクションを返す | 成功 | 8 ms | 別の実行が完了するまで待ってから再実行してください。必要ならログフォルダで直近の実行状況を確認してください。 |
| 14 | Get-ProfileConfiguration は日本語勤怠プロファイルを読み込む | 成功 | 8 ms | 日本語勤怠管理表プロファイル / 35 |
| 15 | Get-ProfileConfiguration は日本語売上日報プロファイルを読み込む | 成功 | 6 ms | 日本語売上日報プロファイル / 12 |
| 16 | Get-ProfileConfiguration は日本語在庫一覧プロファイルを読み込む | 成功 | 5 ms | 日本語在庫一覧プロファイル / 10 |
| 17 | Get-ProfileConfiguration は日本語問い合わせ管理表プロファイルを読み込む | 成功 | 9 ms | 日本語問い合わせ管理表プロファイル / 9 |
| 18 | Get-ProfileConfiguration は生データ転記サンプルプロファイルを読み込む | 成功 | 34 ms | 生データ転記サンプルプロファイル / 30 |
| 19 | Get-ProfileConfiguration は危険な ProfileName を拒否する | 成功 | 13 ms | 危険な ProfileName を拒否 |
| 20 | Get-ProfileConfiguration は外部 ProfilePath を既定拒否し明示許可で読み込む | 成功 | 49 ms | external_profile |
| 21 | TemplateBuilder VBA は V2 テンプレート生成マクロを持つ | 成功 | 3 ms | Review / BuildPDF2ExcelV2 / 5 シート / 重複除去を確認 |
| 22 | テンプレート再作成手順は配置先を明記する | 成功 | 8 ms | README / user-manual / handoff README の配置先明記を確認 |
| 23 | run_pdf2excel.ps1 は Secure モードでローカル runtime を使う | 成功 | 41 ms | SecurityMode / LogLevel / LOCALAPPDATA runtime/logs / path mask / template integrity / KeepInput 無効化を確認 |
| 24 | run_pdf2excel_menu.ps1 はプロファイル選択と環境チェック導線を持つ | 成功 | 2 ms | profile select / environment check / run report / current profile を確認 |
| 25 | V2 ラッパーは Secure モードを渡し RemoteSigned で起動する | 成功 | 4 ms | SecurityMode Secure / RemoteSigned / Bypass 除去を確認 |
| 26 | V2 BAT は ASCII かつ RemoteSigned で起動する | 成功 | 33 ms | ASCII / RemoteSigned / Bypass 除去を確認 |
| 27 | 正式運用 BAT は V2 Secure を起動し RemoteSigned に統一されている | 成功 | 9 ms | run_pdf2excel.bat は V2 Secure、legacy 導線も RemoteSigned を確認 |
| 28 | テストとビルド導線に Bypass が残っていない | 成功 | 4 ms | tests / build 呼び出しの Bypass 除去を確認 |
| 29 | V2 BAT とメニューは ForceMenu 導線を持つ | 成功 | 12 ms | ForceMenu 導線を確認 |
| 30 | 共通メニューは共有パス制約とプロファイル雛形導線を持つ | 成功 | 2 ms | 共有パス制約 / LOCALAPPDATA 必須 / プロファイル雛形 / メニュー番号を確認 |
| 31 | 共通メニューは完了メッセージを版別に分ける | 成功 | 5 ms | V1/V2 完了メッセージと override 分岐を確認 |
| 32 | handoff ビルドは VBA モジュールを同梱する | 成功 | 2 ms | template\vba / sjis ミラー同期 / new_profile_scaffold.ps1 / TemplateBuilder.bas / PDF2ExcelMacros.bas を確認 |
| 33 | 生成済み V2 handoff は配布元ソースと同期している | 成功 | 47 ms | README / BAT / scripts / profile の handoff 同期を確認 |
| 34 | new_profile_scaffold は v1/v2 雛形を生成できる | 成功 | 1062 ms |  VER1 のプロファイル雛形を作成しました。 内部ID: sample_v1 表示名: テストV1 保存先: C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v1.json  次に確認してください: - expectedColumns / headerRowsToSkip / targetRowCount  VER2 のプロファイル雛形を作成しました。 内部ID: sample_v2 表示名: テストV2 保存先: C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v2.json  次に確認してください: - expectedColumns / headerRowsToSkip / targetRowCount - multiPageMergeMode - preferredTableNameContains - normalizedTimeColumns - reviewPersonColumn / reviewSiteColumn / reviewInTimeColumn / reviewOutTimeColumn v1/v2 雛形生成を確認 |
| 35 | new_profile_scaffold は v2 Wizard で主要設定を生成できる | 成功 | 627 ms | sample_v2_wizard テストV2ウィザード C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v2_wizard.json ウィザード説明 28 2 6 Y sameHeader 元ファイル名 項目 勤怠,現場 3 5 6 7 6:正規化入場1,7:正規化退場1  作成内容の確認   VersionMode               : v2   ProfileName               : sample_v2_wizard   DisplayName               : テストV2ウィザード   OutputPath                : C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v2_wizard.json   Description               : ウィザード説明   ExpectedColumns           : 28   HeaderRowsToSkip          : 2   TargetRowCount            : 6   AllowMoreColumns          : True   SourceFileColumnName      : 元ファイル名   DataColumnPrefix          : 項目   PreferredTableNameContains: 勤怠, 現場   MultiPageMergeMode        : sameHeader   ReviewPersonColumn        : 3   ReviewSiteColumn          : 5   ReviewInTimeColumn        : 6   ReviewOutTimeColumn       : 7   NormalizedTimeColumns     : 6:正規化入場1, 7:正規化退場1  Y  VER2 のプロファイル雛形を作成しました。 内部ID: sample_v2_wizard 表示名: テストV2ウィザード 保存先: C:\Work_Codex\PDF2Excel\tests\results\profile-scaffold\sample_v2_wizard.json  次に確認してください: - expectedColumns / headerRowsToSkip / targetRowCount - multiPageMergeMode - preferredTableNameContains - normalizedTimeColumns - reviewPersonColumn / reviewSiteColumn / reviewInTimeColumn / reviewOutTimeColumn v2 Wizard 雛形生成を確認 |
| 36 | build_handoff_package は V2 既定生成と legacy v1 参照を持つ | 成功 | 2 ms | TargetVersion 既定値 v2 と legacy v1 配布元を確認 |
| 37 | build_sample_pdfs は Wizard 体験用 V2 サンプルを生成する | 成功 | 16 ms | profile_wizard_demo sample generation ready |
| 38 | run_integration_tests は Suite 指定で smoke full legacy を切り替えられる | 成功 | 9 ms | Suite switch / report summary を確認 |
| 39 | run.lock は匿名化され、Secure cleanup 失敗時の警告を持つ | 成功 | 2 ms | run.lock 匿名化 / cleanup 警告強化を確認 |
| 40 | Write-RunReport は既存ファイルを原子的に置き換える | 成功 | 81 ms | RunReport の原子的置換を確認 |
| 41 | Append-RunHistory は既存 CSV を原子的に追記する | 成功 | 54 ms | run-history の原子的追記を確認 |
| 42 | Write-EnvironmentCheckReport は既存レポートを原子的に置き換える | 成功 | 31 ms | 環境チェックレポートの原子的置換を確認 |
| 43 | Get-RunLockState は stale lock を識別する | 成功 | 62 ms | stale lock の識別を確認 |
| 44 | Get-RunLockState は壊れた lock を内容読取不可として保持する | 成功 | 16 ms | 壊れた lock の unreadable 判定を確認 |
| 45 | Acquire-RunLock は放棄 mutex 回復の catch 分岐を持つ | 成功 | 2 ms | 放棄 mutex 回復分岐を確認 |
| 46 | テンプレート整合性マニフェストは必要ファイルを持つ | 成功 | 11 ms | template-integrity.json の主要エントリを確認 |
| 47 | 日本語コンソール出力スクリプトは UTF-8 初期化と保存形式を持つ | 成功 | 52 ms | run/build/scaffold scripts の UTF-8 初期化と build_handoff_package.ps1 の BOM を確認 |
| 48 | Write-Log は INFO で詳細パスを伏せ DEBUG を抑止する | 成功 | 11 ms | INFO では path mask / DEBUG 抑止を確認 |
| 49 | Write-Log は DEBUG 指定時だけ詳細パスを出力する | 成功 | 6 ms | DEBUG では詳細パスを確認 |
| 50 | ログ要約関数は件数と代表ファイル名を返す | 成功 | 15 ms | 4 件 (alpha.pdf, beta.pdf, gamma.pdf ほか 1 件) |
| 51 | run_pdf2excel.ps1 は待機メッセージを表示する | 成功 | 2 ms | 待機メッセージを確認 |
| 52 | Normalize-TimeText は全角コロンを半角時刻へ正規化する | 成功 | 28 ms | 08:00 / 480 |
| 53 | Normalize-TimeText は時分表記と空白込みを正規化する | 成功 | 1 ms | 09:15 / 555 |
| 54 | Normalize-TimeText は時のみ表記を 00 分補完する | 成功 | 1 ms | 18:00 / 1080 |
| 55 | Normalize-TimeText は Excel 時刻比率の文字列も正規化する | 成功 | 23 ms | 09:15 / 555 |
| 56 | Normalize-TimeText は 24:00 を有効時刻として扱う | 成功 | 1 ms | 24:00 / 1440 |
| 57 | Normalize-TimeText は 24:30 を範囲外扱いにする | 成功 | 2 ms | INVALID / 時刻の範囲外です。 |
| 58 | Normalize-TimeText は 整数文字列 1 を 01:00 として扱う | 成功 | 1 ms | 01:00 / 60 |
| 59 | Normalize-TimeText は 24:00:00 を 24:00 として扱う | 成功 | 6 ms | 24:00 / 1440 |
| 60 | Get-ResultOutputColumnNames は V2 で正規化列を追加する | 成功 | 18 ms | 正規化入場2_分,正規化退場2,正規化退場2_分,時刻正規化状態,時刻確認メモ |
| 61 | Get-NormalizedTimeColumnDefinitions は Review raw 列名を返す | 成功 | 13 ms | 正規化入場1_raw,正規化退場1_raw |
| 62 | Get-ReviewReasonCategories は HEADER_MISMATCH を先頭に重複なく返す | 成功 | 77 ms | HEADER_MISMATCH,TIME_MULTI,TIME_MISSING,TIME_INVALID |
| 63 | Get-StagingQueryFormula は V2 で canonical sameHeader と曖昧分離を考慮する | 成功 | 39 ms | canonical sameHeader / ambiguity split / review raw / reason categories |
| 64 | Get-ReviewQueryFormulaV2 は ReasonCategory を展開対象に含める | 成功 | 15 ms | ReasonCategory expansion ready |
| 65 | run_pdf2excel.bat は ASCII のみで構成される | 成功 | 11 ms | ASCII のみを確認 |
| 66 | run_pdf2excel_v1.bat と run_pdf2excel_v2.bat は ASCII のみで構成される | 成功 | 13 ms | 版別 BAT の ASCII を確認 |
| 67 | run_pdf2excel_menu.ps1 は UTF-8 BOM で保存される | 成功 | 1 ms | UTF-8 BOM を確認 |
| 68 | 版別メニュー PowerShell は UTF-8 BOM で保存される | 成功 | 1 ms | 版別メニューの UTF-8 BOM を確認 |
