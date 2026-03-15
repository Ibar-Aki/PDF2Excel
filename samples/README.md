# サンプルデータ

- 作成日: 2026-03-15 01:10 JST
- 作成者: Codex (GPT-5)

`samples` フォルダには、PDF2Excel の動作確認に使えるサンプル帳票を置いています。

## 日本語勤怠管理表サンプル

- PDF: `pdf/attendance_jp/2026年03月_勤怠管理表.pdf`
- PDF: `pdf/attendance_jp/2026年04月_勤怠管理表.pdf`
- 元Excel: `source/attendance_jp/2026年03月_勤怠管理表.xlsx`
- 元Excel: `source/attendance_jp/2026年04月_勤怠管理表.xlsx`
- 推奨プロファイル: `../config/profiles/attendance_monthly_jp.json`

## 日本語売上日報サンプル

- PDF: `pdf/sales_daily_jp/2026-03-15_売上日報_東京店.pdf`
- PDF: `pdf/sales_daily_jp/2026-03-16_売上日報_横浜店.pdf`
- 元Excel: `source/sales_daily_jp/2026-03-15_売上日報_東京店.xlsx`
- 元Excel: `source/sales_daily_jp/2026-03-16_売上日報_横浜店.xlsx`
- 推奨プロファイル: `../config/profiles/sales_daily_jp.json`

## 日本語在庫一覧サンプル

- PDF: `pdf/inventory_jp/春季_在庫一覧_倉庫A.pdf`
- PDF: `pdf/inventory_jp/春季_在庫一覧_倉庫B.pdf`
- 元Excel: `source/inventory_jp/春季_在庫一覧_倉庫A.xlsx`
- 元Excel: `source/inventory_jp/春季_在庫一覧_倉庫B.xlsx`
- 推奨プロファイル: `../config/profiles/inventory_list_jp.json`

## 日本語問い合わせ管理表サンプル

- PDF: `pdf/inquiry_jp/2026年03月_問い合わせ管理表_第1週.pdf`
- PDF: `pdf/inquiry_jp/2026年03月_問い合わせ管理表_第2週.pdf`
- 元Excel: `source/inquiry_jp/2026年03月_問い合わせ管理表_第1週.xlsx`
- 元Excel: `source/inquiry_jp/2026年03月_問い合わせ管理表_第2週.xlsx`
- 推奨プロファイル: `../config/profiles/inquiry_weekly_jp.json`

## 建設現場転記 PoC サンプル

- PDF: `pdf/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC.pdf`
- 元Excel: `source/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC.xlsx`
- 推奨プロファイル: `../config/profiles/construction_transfer_poc.json`
- 特徴:
  - タイトル 3 行 + 4 行目ヘッダー
  - 30 列
  - セル内改行あり
  - 時刻表記ゆれあり
  - 現場名と氏名の微妙な揺れあり

## 再生成方法

次のコマンドで、日本語サンプル帳票一式を再生成できます。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_sample_pdfs.ps1
```

## 試し方

1. 試したい帳票タイプの PDF フォルダを対象にします。
2. 対応するプロファイル JSON を指定して変換します。
3. `Result` シートで日本語の項目名と値が正しく入っていることを確認します。
