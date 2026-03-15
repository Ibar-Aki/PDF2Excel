# PDF2Excel ユーザーマニュアル

- 作成日: 2026-03-13 00:05 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-15

補助資料:

- 文書一覧: [index.md](index.md)
- 構成ガイド: [project-layout.md](project-layout.md)
- 障害対応: [troubleshooting.md](troubleshooting.md)
- サンプルPDF: `samples/v1/pdf`, `samples/v2/pdf`
- サンプル一覧: [../samples/README.md](../samples/README.md)

## 1. このツールの概要

このツールは、PDF に入っている表をまとめて Excel に変換するためのものです。  
複数の PDF を一括で読み込み、1つの `xlsx` に集約します。

出力される Excel は、次の形に統一されます。

- A列: PDF ファイル名
- B列以降: プロファイルに応じた表データ列
- `Summary` シート: 成功件数、失敗件数、PDFごとの内訳
- `Errors` シート: 壊れた PDF や列数不一致などの失敗情報

`VER2` では、これに加えて `Review` シートが出ます。  
`Review` には、確認が必要な raw 行が `元ファイル名 / ページ / 氏名 raw / 現場 raw / 入場 raw / 退場 raw / 確認要理由` で並びます。
さらに `Result` の末尾に `正規化入場*` / `正規化退場*` / `*_分` / `時刻正規化状態` / `時刻確認メモ` が追加され、`08：00` や `9時 15分`、Excel 時刻比率文字列を分析向けにそろえやすくしています。

向いている用途:

- 毎月・毎週届く同じ帳票レイアウトの PDF をまとめて Excel 化したい
- 50件前後の PDF を1件ずつコピペせず一括処理したい
- 誰でも同じ手順で実行できるようにしたい

前提条件:

- Windows 上で Excel(M365) デスクトップ版が使える
- PDF が画像ではなく、文字を選択できるテキスト PDF である
- 対象の表はおおむね同じ列構成である

## 2. まずこれだけ見れば使える最短手順

1. 標準変換なら [run_pdf2excel_v1.bat](../run_pdf2excel_v1.bat)、建設現場 raw 転記なら [run_pdf2excel_v2.bat](../run_pdf2excel_v2.bat) をダブルクリックします。
2. 表示されたメニューで `1` を押します。
3. 変換したい PDF を複数選びます。
4. 出力する Excel ファイルの保存先を選びます。
5. 実行前チェックの内容を確認して `Y` を押します。
6. 完了後、保存した `xlsx` を開きます。
7. `Summary` シートで全体件数を確認します。
8. `Result` シートを確認します。
9. 失敗した PDF がないか `Errors` シートも確認します。

迷ったときは、まずこのやり方で十分です。

## 2-1. 今回の安全対策

- 同時に 2 回以上は実行できません。
- ほかの実行が動いているときは、その場で停止します。
- 実際の変換は毎回専用の一時 staging で行います。
- `input` に残っている過去PDFは、`-KeepInput` を使っても今回の変換には混ざりません。

## 3. BAT メニューの意味

### [1] PDFファイルを選んで変換

- もっともかんたんな使い方です。
- 1つの選択ダイアログから複数の PDF を選べます。
- 毎回対象ファイルが変わる場合に向いています。

### [2] PDFフォルダを選んで変換

- 1つのフォルダに PDF をまとめてある場合に向いています。
- フォルダ選択ダイアログが開きます。
- 毎回同じフォルダを処理する運用で使いやすい方法です。

### [3] 使い方マニュアルを開く

- このマニュアルを開きます。
- 迷ったときの確認用です。

### [4] 出力フォルダを開く

- `output` フォルダを開きます。
- 過去の出力ファイルを見たいときに使います。

### [5] プロファイルフォルダを開く

- `config\profiles` フォルダを開きます。
- 帳票プロファイルを確認、複製、編集したいときに使います。
- 日本語の月次勤怠管理表を試す場合は `attendance_monthly_jp.json` を確認してください。
- 日本語の売上日報、在庫一覧、問い合わせ管理表も、それぞれ専用プロファイルを同梱しています。

### [6] ログフォルダを開く

- `logs` フォルダを開きます。
- 直近の実行ログを確認したいときに使います。

### [7] 終了

- 何もせず終了します。

## 4. 変換前の準備

変換前に次を確認すると失敗が減ります。

### 4-0. 実行前チェックとは

変換開始前に、次の内容が画面に表示されます。

- 対象 PDF 数
- 出力ファイル
- 上書き有無
- 使用プロファイル
- 想定列数
- ヘッダー除外行数
- 想定行数
- 代表ファイル名

内容に誤りがあれば `N` で中止してください。  
自動実行やテストでは `-NoConfirm` を付けると確認入力を省略できます。

### 4-1. PDF の条件

- PDF を開いて文字をドラッグ選択できること
- 対象の表が、使用するプロファイルの想定列数に近いこと
- 似たレイアウトの帳票をまとめて処理すること

### 4-2. ファイル名の条件

- 同じ名前の PDF を同時に処理しないこと
- 例:
  - `A\report.pdf`
  - `B\report.pdf`

上記のように同名だと、どちらがどの結果かわからなくなるため、ツールはエラーで停止します。

### 4-4. `KeepInput` の意味

`-KeepInput` は、`input` フォルダに保管済みのPDFを消さずに残すためのオプションです。  
ただし、実際の変換は毎回専用の staging フォルダで実行されます。  
そのため、以前のPDFが今回の `Result` に混ざることはありません。

### 4-3. 事前に分けておくとよいケース

- レイアウトが大きく違う PDF
- 画像PDFやスキャンPDF
- 30列を大きく超える表

これらは別グループで扱うか、対象外として切り分けてください。

## 5. 実際の操作手順

### パターンA: ファイルを選んで変換する

1. `run_pdf2excel.bat` を起動します。
2. `1` を押します。
3. 変換したい PDF を選びます。
4. `開く` を押します。
5. 保存先の `xlsx` を選びます。
6. 実行前チェックを確認して `Y` を押します。
7. 処理完了を待ちます。
8. 保存した `xlsx` を開いて確認します。

### パターンB: フォルダごと変換する

1. `run_pdf2excel.bat` を起動します。
2. `2` を押します。
3. PDF が入っているフォルダを選びます。
4. 保存先の `xlsx` を選びます。
5. 実行前チェックを確認して `Y` を押します。
6. 処理完了を待ちます。
7. 保存した `xlsx` を開いて確認します。

### パターンC: PowerShell から実行する

ファイル選択や保存先選択をダイアログで行う場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -SelectInputFolder -PromptForOutputFile
```

入力フォルダと出力先を直接指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx
```

ファイルを個別指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFiles "C:\PDF\a.pdf,C:\Other\b.pdf" -OutputFile C:\Work\result.xlsx
```

完成した Excel を自動で開きたい場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx -OpenOutput
```

プロファイルを名前で指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -ProfileName default -OutputFile C:\Work\result.xlsx
```

プロファイルを JSON ファイルで直接指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel_v1.ps1 -InputFolder C:\Work\pdf -ProfilePath C:\Work\custom-profile.json -OutputFile C:\Work\result.xlsx
```

## 6. 帳票プロファイルの使い方

帳票プロファイルは、帳票ごとの抽出条件をまとめた JSON です。  
既定では `config/profiles/default.json` を使います。  
日本語サンプル向けには `attendance_monthly_jp.json`、`sales_daily_jp.json`、`inventory_list_jp.json`、`inquiry_weekly_jp.json` も用意しています。  

プロファイルで主に調整する項目:

- `expectedColumns`
  - 想定列数です。
- `headerRowsToSkip`
  - 先頭から除外するヘッダー行数です。
- `targetRowCount`
  - 表候補を選ぶときの目安行数です。
- `allowMoreColumns`
  - 想定列数を超える表を許可するかどうかです。
- `preferredTableKinds`
  - 優先したい表種別です。
- `preferredTableNameContains`
  - 表名に含まれていると優先する文字列です。
- `preferredTableIdContains`
  - 表 ID に含まれていると優先する文字列です。

おすすめの始め方:

1. `default.json` をコピーして別名にします。
2. `expectedColumns` と `headerRowsToSkip` だけ先に調整します。
3. 少数の PDF で試します。
4. `Summary` と `Errors` を見て、調整が効いているか確認します。

日本語勤怠管理表を試す場合:

1. `attendance_monthly_jp.json` をそのまま使って試します。
2. 列数が 35 列前後、氏名や日別勤怠が横に並ぶ帳票に向いています。
3. 6 人前後の月次勤怠表でまず少数テストしてから本番投入してください。
4. すぐ試す場合は `samples/pdf/attendance_jp` のサンプル PDF を使えます。

日本語売上日報を試す場合:

1. `sales_daily_jp.json` を使います。
2. `samples/pdf/sales_daily_jp` のサンプル PDF を使えます。
3. 店舗名、商品名、売上金額が `Result` に入るか確認します。

日本語在庫一覧を試す場合:

1. `inventory_list_jp.json` を使います。
2. `samples/pdf/inventory_jp` のサンプル PDF を使えます。
3. 品番、品名、倉庫、在庫数が `Result` に入るか確認します。

日本語問い合わせ管理表を試す場合:

1. `inquiry_weekly_jp.json` を使います。
2. `samples/pdf/inquiry_jp` のサンプル PDF を使えます。
3. 顧客名、件名、状態が `Result` に入るか確認します。

## 7. 実行中に表示される内容の見方

PowerShell 実行時には、主に次の情報が表示されます。

- 実行前チェック
- 対象 PDF 数
- 読み込み準備の完了
- Excel 起動中
- Power Query 設定中
- Result / Errors / Summary シートへの読込中
- 実行結果の要約

最後に表示される `実行結果` の見方:

- `対象PDF数`: 読み込んだ PDF の数
- `成功PDF数`: 正常に変換できた PDF の数
- `失敗PDF数`: `Errors` シートへ出た PDF の数
- `取込データ行数`: `Result` シートへ入った明細行数
- `エラー件数`: `Errors` シートに出た件数
- `処理時間(秒)`: 実行にかかった時間
- `使用プロファイル`: 今回使った帳票プロファイル
- `出力ファイル`: 作成された `xlsx`
- `ログファイル`: 詳細な実行ログ

## 8. 出力される Excel の見方

### Control シート

- 実行の基本情報が入ります。
- 主な確認項目:
  - 入力フォルダ
  - 出力ファイル
  - ログファイル
  - 最終実行日時
  - 状態
  - 対象PDF数
  - 取込データ行数
  - エラー件数

初見の人は、まずこのシートを見ると全体像がわかります。
`入力フォルダ` には、今回の変換に使った専用 staging フォルダが記録されます。

### Summary シート

- 実行結果の総括が入ります。
- 主な確認項目:
  - 対象PDF数
  - 成功PDF数
  - 失敗PDF数
  - 取込データ行数
  - エラー件数
  - 処理時間
  - PDF ごとの取込件数
  - エラー分類別件数

### Result シート

- 1行目は見出しです。
- 2行目以降が変換結果です。
- A列に元 PDF 名が入ります。
- B列以降にプロファイルに応じた表データが入ります。
- `VER2` では末尾に `正規化入場*` / `正規化退場*` / `*_分` が追加されます。
- `*_分` は 0:00 からの分換算で、後続の集計や比較に使えます。
- `時刻正規化状態` が `OK` なら正規化済み、`要確認` なら元セルを見直してください。

確認ポイント:

- PDF ごとの行が A列のファイル名で区別できているか
- 空欄がずれて別列に入っていないか
- 想定件数と行数が近いか

### Review シート

- `VER2` のみです。
- 確認が必要な raw 行を、元ファイル名、ページ、氏名 raw、現場 raw、入場 raw、退場 raw、理由で一覧します。
- あわせて `正規化入場` / `正規化退場` / `*_分` / `時刻正規化状態` / `時刻確認メモ` を出します。
- `Result` 側は全件、`Review` 側は確認優先行という役割分担です。

### Errors シート

- 変換に失敗した PDF がある場合に使います。
- 主な例:
  - PDF が壊れている
  - 表を見つけられない
  - プロファイルの想定列数と合わない

`Errors` シートには次の情報が入ります。

- `ErrorCode`
- `ErrorCategory`
- `UserMessage`
- `TechnicalDetail`

`Errors` シートに何もなければ、変換ロジック上の異常は起きていません。

## 9. 失敗したときの見方

確認の順番は次の通りです。

1. `Errors` シートを見る
2. `Control` シートの状態と件数を見る
3. `logs` フォルダの最新ログを見る
4. 元 PDF を開いて文字選択できるかを見る
5. レイアウトが他の PDF と大きく違わないかを見る

主な原因と対処:

| 症状 | よくある原因 | 対処 |
| --- | --- | --- |
| 何も出力されない | PDF 未選択、入力フォルダ空、Excel 起動失敗 | PDF の選択と Excel 利用可否を確認 |
| `Errors` に多く出る | レイアウト差が大きい、またはプロファイル不一致 | 似た帳票ごとに分けるか、プロファイルを見直す |
| 列がずれる | 元PDFの表認識が不安定 | 実PDFを数件確認し、対象グループを絞る |
| 同名PDFエラー | 別フォルダでファイル名が重複 | 事前に片方の名前を変更 |
| 壊れたPDF | PDF が不完全 | 元ファイルを再取得 |
| 実行前チェックの内容が違う | 入力元、保存先、プロファイル指定ミス | いったん `N` で中止して選び直す |

## 10. おすすめの運用方法

### パターンA: 毎回ばらばらのPDFを処理したい

- BAT を起動
- `1` を選ぶ
- 必要な PDF だけ選択

### パターンB: 毎回同じフォルダを処理したい

- BAT を起動
- `2` を選ぶ
- PDF が入っているフォルダを入力

### パターンC: PowerShell から明示的に実行したい

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx
```

## 11. よくある質問

### Q. どんな PDF でも読めますか？

いいえ。  
OCR 前提の画像 PDF ではなく、文字を選択できるテキスト PDF を前提にしています。

### Q. PDF のレイアウトが少し違っても使えますか？

少しの差なら読める場合があります。  
ただし、列構成が大きく違うと `Errors` シートへ出る可能性があります。必要に応じて別プロファイルを作成してください。

### Q. 同じ名前の PDF を別フォルダから一緒に選べますか？

できません。  
同名ファイルは衝突するため、エラーで止まります。先にファイル名を変えてください。

### Q. `input` フォルダをそのまま指定しても大丈夫ですか？

大丈夫です。  
現在の版では自己削除しないように修正済みです。

### Q. 帳票が 30 列ではない場合も使えますか？

使えます。  
`config\profiles` にあるプロファイルを複製し、`expectedColumns` を調整してください。

## 12. ファイルとログの場所

- 実行ログ:
  - `logs`
  - 古いログは 30 日超または 200 件超で自動整理されます。
- 出力された Excel:
- `output`
- プロファイル:
- `config/profiles`
- サンプル PDF:
- `samples/pdf`
- テンプレート:
- `template/PDF2Excel_Converter.xlsm`
- テスト結果レポート:
- [test-report.md](../reports/test-report.md)

## 13. テスト済みの内容

このツールは次の観点で統合テスト済みです。

- PowerShell 実行
- BAT 実行
- 実行前チェックの自動スキップ
- プロファイル切替
- `input` 自己参照
- 壊れた PDF 混在
- 深い出力先フォルダ
- 一時ファイル清掃
- Excel プロセス残留なし

詳しい結果は [test-report.md](../reports/test-report.md) を参照してください。

## 14. 最後に

まずは少数の実PDFで一度変換し、`Result` と `Errors` の出方を確認してください。  
問題なければ、そのまま 50 件前後の一括処理へ広げるのが安全です。
