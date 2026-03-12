# PDF2Excel ユーザーマニュアル

- 作成日: 2026-03-13 00:05 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-13

補助資料:

- 構成ガイド: [project-layout.md](C:\Work_Codex\PDF2Excel\docs\project-layout.md)
- サンプルPDF: `C:\Work_Codex\PDF2Excel\samples\pdf`

## 1. このツールの概要

このツールは、PDF に入っている表をまとめて Excel に変換するためのものです。  
複数の PDF を一括で読み込み、1つの `xlsx` に集約します。

出力される Excel は、次の形に統一されます。

- A列: PDF ファイル名
- B列〜AE列: 表の 30 列分
- `Errors` シート: 壊れた PDF や列数不一致などの失敗情報

向いている用途:

- 毎月・毎週届く同じ帳票レイアウトの PDF をまとめて Excel 化したい
- 50件前後の PDF を1件ずつコピペせず一括処理したい
- 誰でも同じ手順で実行できるようにしたい

前提条件:

- Windows 上で Excel(M365) デスクトップ版が使える
- PDF が画像ではなく、文字を選択できるテキスト PDF である
- 対象の表はおおむね同じ列構成である

## 2. まずこれだけ見れば使える最短手順

1. [run_pdf2excel.bat](C:\Work_Codex\PDF2Excel\run_pdf2excel.bat) をダブルクリックします。
2. 表示されたメニューで `1` を押します。
3. 変換したい PDF を複数選びます。
4. 出力する Excel ファイルの保存先を選びます。
5. 完了後、保存した `xlsx` を開きます。
6. `Result` シートを確認します。
7. 失敗した PDF がないか `Errors` シートも確認します。

迷ったときは、まずこのやり方で十分です。

## 3. BAT メニューの意味

### [1] PDFファイルを選んで変換

- もっともかんたんな使い方です。
- 複数のフォルダにある PDF を一度に選べます。
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

### [5] 終了

- 何もせず終了します。

## 4. 変換前の準備

変換前に次を確認すると失敗が減ります。

### 4-1. PDF の条件

- PDF を開いて文字をドラッグ選択できること
- 対象の表がおおむね 30 列以内であること
- 似たレイアウトの帳票をまとめて処理すること

### 4-2. ファイル名の条件

- 同じ名前の PDF を同時に処理しないこと
- 例:
  - `A\report.pdf`
  - `B\report.pdf`

上記のように同名だと、どちらがどの結果かわからなくなるため、ツールはエラーで停止します。

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
6. 処理完了を待ちます。
7. 保存した `xlsx` を開いて確認します。

### パターンB: フォルダごと変換する

1. `run_pdf2excel.bat` を起動します。
2. `2` を押します。
3. PDF が入っているフォルダを選びます。
4. 保存先の `xlsx` を選びます。
5. 処理完了を待ちます。
6. 保存した `xlsx` を開いて確認します。

### パターンC: PowerShell から実行する

ファイル選択や保存先選択をダイアログで行う場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -SelectInputFolder -PromptForOutputFile
```

入力フォルダと出力先を直接指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx
```

ファイルを個別指定する場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFiles "C:\PDF\a.pdf,C:\Other\b.pdf" -OutputFile C:\Work\result.xlsx
```

完成した Excel を自動で開きたい場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run_pdf2excel.ps1 -InputFolder C:\Work\pdf -OutputFile C:\Work\result.xlsx -OpenOutput
```

## 6. 実行中に表示される内容の見方

PowerShell 実行時には、主に次の情報が表示されます。

- 対象 PDF 数
- 読み込み準備の完了
- Excel 起動中
- Power Query 設定中
- Result / Errors シートへの読込中
- 実行結果の要約

最後に表示される `実行結果` の見方:

- `対象PDF数`: 読み込んだ PDF の数
- `取込データ行数`: `Result` シートへ入った明細行数
- `エラー件数`: `Errors` シートに出た件数
- `出力ファイル`: 作成された `xlsx`
- `ログファイル`: 詳細な実行ログ

## 7. 出力される Excel の見方

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

### Result シート

- 1行目は見出しです。
- 2行目以降が変換結果です。
- A列に元 PDF 名が入ります。
- B列〜AE列に表データが入ります。

確認ポイント:

- PDF ごとの行が A列のファイル名で区別できているか
- 空欄がずれて別列に入っていないか
- 想定件数と行数が近いか

### Errors シート

- 変換に失敗した PDF がある場合に使います。
- 主な例:
  - PDF が壊れている
  - 表を見つけられない
  - 30列を超えていて対象外

`Errors` シートに何もなければ、変換ロジック上の異常は起きていません。

## 8. 失敗したときの見方

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
| `Errors` に多く出る | レイアウト差が大きい | 似た帳票ごとに分けて実行 |
| 列がずれる | 元PDFの表認識が不安定 | 実PDFを数件確認し、対象グループを絞る |
| 同名PDFエラー | 別フォルダでファイル名が重複 | 事前に片方の名前を変更 |
| 壊れたPDF | PDF が不完全 | 元ファイルを再取得 |

## 9. おすすめの運用方法

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

## 10. よくある質問

### Q. どんな PDF でも読めますか？

いいえ。  
OCR 前提の画像 PDF ではなく、文字を選択できるテキスト PDF を前提にしています。

### Q. PDF のレイアウトが少し違っても使えますか？

少しの差なら読める場合があります。  
ただし、列構成が大きく違うと `Errors` シートへ出る可能性があります。

### Q. 同じ名前の PDF を別フォルダから一緒に選べますか？

できません。  
同名ファイルは衝突するため、エラーで止まります。先にファイル名を変えてください。

### Q. `input` フォルダをそのまま指定しても大丈夫ですか？

大丈夫です。  
現在の版では自己削除しないように修正済みです。

## 11. ファイルとログの場所

- 実行ログ:
  - `C:\Work_Codex\PDF2Excel\logs`
- 出力された Excel:
  - `C:\Work_Codex\PDF2Excel\output`
- サンプル PDF:
  - `C:\Work_Codex\PDF2Excel\samples\pdf`
- テンプレート:
  - `C:\Work_Codex\PDF2Excel\template\PDF2Excel_Converter.xlsm`
- テスト結果レポート:
  - [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md)

## 12. テスト済みの内容

このツールは次の観点で統合テスト済みです。

- PowerShell 実行
- BAT 実行
- `input` 自己参照
- 壊れた PDF 混在
- 深い出力先フォルダ
- 一時ファイル清掃
- Excel プロセス残留なし

詳しい結果は [test-report.md](C:\Work_Codex\PDF2Excel\reports\test-report.md) を参照してください。

## 13. 最後に

まずは少数の実PDFで一度変換し、`Result` と `Errors` の出方を確認してください。  
問題なければ、そのまま 50 件前後の一括処理へ広げるのが安全です。
