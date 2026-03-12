# PDF2Excel テストレポート

- 作成日: 2026-03-13 00:41 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-13

## 概要

- 実施日時: 2026-03-13 00:41:10 JST - 2026-03-13 00:43:29 JST
- 対象環境: Windows / PowerShell 5.1.26100.7920 / Excel(M365) COM
- 対象URLまたは対象機能: ローカル PowerShell / BAT / Excel(M365) による PDF2Excel 一括変換
- 結果概要: 12件成功 / 0件失敗
- 所要時間または主要な応答時間: 全体 139 秒
- エラー有無: なし
- 失敗時の原因推定: 該当なし

## 実行シナリオ

| No | シナリオ | 結果 | 所要時間 | 補足 |
| --- | --- | --- | --- | --- |
| 1 | Template build | PASS | 4845 ms | Template created: C:\Work_Codex\PDF2Excel\scripts\..\template\PDF2Excel_Converter.xlsm Template updated at 03/12/2026 15:41:19 |
| 2 | PowerShell conversion | PASS | 14500 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:41:23 [INFO] 入力 PDF を確認しています。 2026-03-13 00:41:23 [INFO] 対象 PDF 数: 2 2026-03-13 00:41:23 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf, C:\Work_Codex\PDF2Excel\input\valid_b.pdf 2026-03-13 00:41:25 [INFO] Excel を起動しています。 2026-03-13 00:41:27 [INFO] Power Query を設定しています。 2026-03-13 00:41:29 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:41:34 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:41:34 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\powershell_success.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 8   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\powershell_success.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004123.log  ResultRows=9, ErrorsRows=1 |
| 3 | Single PDF conversion | PASS | 13693 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:41:41 [INFO] 入力 PDF を確認しています。 2026-03-13 00:41:41 [INFO] 対象 PDF 数: 1 2026-03-13 00:41:41 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf 2026-03-13 00:41:43 [INFO] Excel を起動しています。 2026-03-13 00:41:44 [INFO] Power Query を設定しています。 2026-03-13 00:41:46 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:41:50 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:41:51 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\single_pdf_success.xlsx  実行結果   対象PDF数       : 1   取込データ行数   : 4   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\single_pdf_success.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004141.log  Single PDF conversion succeeded |
| 4 | InputFiles conversion | PASS | 13821 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:41:58 [INFO] 入力 PDF を確認しています。 2026-03-13 00:41:58 [INFO] 対象 PDF 数: 2 2026-03-13 00:41:58 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf, C:\Work_Codex\PDF2Excel\input\valid_b.pdf 2026-03-13 00:42:00 [INFO] Excel を起動しています。 2026-03-13 00:42:01 [INFO] Power Query を設定しています。 2026-03-13 00:42:03 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:42:07 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:42:08 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\inputfiles_success.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 8   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\inputfiles_success.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004157.log  InputFiles conversion succeeded |
| 5 | BAT conversion | PASS | 13807 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:42:14 [INFO] 入力 PDF を確認しています。 2026-03-13 00:42:14 [INFO] 対象 PDF 数: 2 2026-03-13 00:42:14 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf, C:\Work_Codex\PDF2Excel\input\valid_b.pdf 2026-03-13 00:42:16 [INFO] Excel を起動しています。 2026-03-13 00:42:18 [INFO] Power Query を設定しています。 2026-03-13 00:42:20 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:42:24 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:42:25 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\bat_success.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 8   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\bat_success.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004214.log   Conversion completed. Check the xlsx file and logs under the output and logs folders if needed. ResultRows=9 |
| 6 | BAT direct no-pause | PASS | 13164 ms | BAT direct execution finished without pause |
| 7 | Input self-reference | PASS | 12956 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:42:44 [INFO] 入力 PDF を確認しています。 2026-03-13 00:42:44 [INFO] 対象 PDF 数: 2 2026-03-13 00:42:44 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf, C:\Work_Codex\PDF2Excel\input\valid_b.pdf 2026-03-13 00:42:46 [INFO] Excel を起動しています。 2026-03-13 00:42:48 [INFO] Power Query を設定しています。 2026-03-13 00:42:50 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:42:54 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:42:55 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\input_self_reference.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 8   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\input_self_reference.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004244.log  RemainingInputFiles=valid_a.pdf,valid_b.pdf |
| 8 | Duplicate filename rejection | PASS | 450 ms | Duplicate PDF names are rejected |
| 9 | Broken PDF handling | PASS | 13767 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:42:59 [INFO] 入力 PDF を確認しています。 2026-03-13 00:42:59 [INFO] 対象 PDF 数: 2 2026-03-13 00:42:59 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\broken.pdf, C:\Work_Codex\PDF2Excel\input\valid_a.pdf 2026-03-13 00:43:01 [INFO] Excel を起動しています。 2026-03-13 00:43:02 [INFO] Power Query を設定しています。 2026-03-13 00:43:04 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:43:09 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:43:09 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\mixed_broken.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 4   エラー件数      : 1   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\mixed_broken.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004259.log  ResultRows=5, ErrorsRows=2 |
| 10 | Nested output path | PASS | 12930 ms |  ==========================================  PDF2Excel - PDF表 一括変換ツール ==========================================  2026-03-13 00:43:16 [INFO] 入力 PDF を確認しています。 2026-03-13 00:43:16 [INFO] 対象 PDF 数: 2 2026-03-13 00:43:16 [INFO] 入力準備が完了しました: C:\Work_Codex\PDF2Excel\input\valid_a.pdf, C:\Work_Codex\PDF2Excel\input\valid_b.pdf 2026-03-13 00:43:18 [INFO] Excel を起動しています。 2026-03-13 00:43:19 [INFO] Power Query を設定しています。 2026-03-13 00:43:21 [INFO] Result / Errors シートへ読み込んでいます。 2026-03-13 00:43:25 [INFO] Excel 出力処理を実行しています。 2026-03-13 00:43:26 [INFO] 処理が完了しました: C:\Work_Codex\PDF2Excel\tests\results\nested\child\output\nested_output.xlsx  実行結果   対象PDF数       : 2   取込データ行数   : 8   エラー件数      : 0   出力ファイル     : C:\Work_Codex\PDF2Excel\tests\results\nested\child\output\nested_output.xlsx   ログファイル     : C:\Work_Codex\PDF2Excel\logs\run_20260313_004316.log  Nested output created |
| 11 | Runtime cleanup | PASS | 2 ms | Runtime directory is clean |
| 12 | No Excel leak | PASS | 4 ms | No Excel process remains |

## 判定

- 主要な正常系、異常系、運用系のシナリオはすべて通過しました。
- 実行後に余分な Excel プロセスが残らないことを確認しました。
- `output/runtime` に一時ファイルが残らず、自動清掃が機能しています。

## 備考

- 抽出精度そのものは `Pdf.Tables` に依存するため、実業務PDFでは列ズレ確認を推奨します。
- テスト用PDFは Excel から生成したテキストPDFを使用しました。
