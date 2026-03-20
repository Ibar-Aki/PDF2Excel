# サンプルデータ

- 作成日: 2026-03-15 01:10 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

`samples` フォルダには、PDF2Excel の動作確認に使えるサンプル帳票を置いています。

- `samples/common`
  - 共通サンプルの正本です。再生成スクリプトはここを更新します。
- `samples/v2`
  - 生データ転記向けサンプルです。
  - 最初に試す順番は [samples/v2/README.md](v2/README.md) を参照してください。
  - 操作手順まで含めた体験ガイドは [01-v2-sample-walkthrough.md](../docs/02-guides/01-v2-sample-walkthrough.md) を参照してください。
- `legacy/v1/samples`
  - `VER1` の互換確認用サンプルです。通常利用者向けの導線からは外しています。

## legacy 用 日本語勤怠管理表サンプル

- PDF: `../legacy/v1/samples/pdf/attendance_jp/2026年03月_勤怠管理表.pdf`
- PDF: `../legacy/v1/samples/pdf/attendance_jp/2026年04月_勤怠管理表.pdf`
- 元Excel: `../legacy/v1/samples/source/attendance_jp/2026年03月_勤怠管理表.xlsx`
- 元Excel: `../legacy/v1/samples/source/attendance_jp/2026年04月_勤怠管理表.xlsx`
- 推奨プロファイル: `../legacy/v1/config/profiles/attendance_monthly_jp.json`

## legacy 用 日本語売上日報サンプル

- PDF: `../legacy/v1/samples/pdf/sales_daily_jp/2026-03-15_売上日報_東京店.pdf`
- PDF: `../legacy/v1/samples/pdf/sales_daily_jp/2026-03-16_売上日報_横浜店.pdf`
- 元Excel: `../legacy/v1/samples/source/sales_daily_jp/2026-03-15_売上日報_東京店.xlsx`
- 元Excel: `../legacy/v1/samples/source/sales_daily_jp/2026-03-16_売上日報_横浜店.xlsx`
- 推奨プロファイル: `../legacy/v1/config/profiles/sales_daily_jp.json`

## legacy 用 日本語在庫一覧サンプル

- PDF: `../legacy/v1/samples/pdf/inventory_jp/春季_在庫一覧_倉庫A.pdf`
- PDF: `../legacy/v1/samples/pdf/inventory_jp/春季_在庫一覧_倉庫B.pdf`
- 元Excel: `../legacy/v1/samples/source/inventory_jp/春季_在庫一覧_倉庫A.xlsx`
- 元Excel: `../legacy/v1/samples/source/inventory_jp/春季_在庫一覧_倉庫B.xlsx`
- 推奨プロファイル: `../legacy/v1/config/profiles/inventory_list_jp.json`

## legacy 用 日本語問い合わせ管理表サンプル

- PDF: `../legacy/v1/samples/pdf/inquiry_jp/2026年03月_問い合わせ管理表_第1週.pdf`
- PDF: `../legacy/v1/samples/pdf/inquiry_jp/2026年03月_問い合わせ管理表_第2週.pdf`
- 元Excel: `../legacy/v1/samples/source/inquiry_jp/2026年03月_問い合わせ管理表_第1週.xlsx`
- 元Excel: `../legacy/v1/samples/source/inquiry_jp/2026年03月_問い合わせ管理表_第2週.xlsx`
- 推奨プロファイル: `../legacy/v1/config/profiles/inquiry_weekly_jp.json`

## 生データ転記サンプル

- PDF: `common/pdf/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC.pdf`
- PDF: `common/pdf/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf`
- PDF: `common/pdf/construction_transfer_poc/2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf`
- PDF: `common/pdf/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf`
- PDF: `common/pdf/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf`
- 元Excel: `common/source/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC.xlsx`
- 元Excel: `common/source/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_2ページ同一列.xlsx`
- 元Excel: `common/source/construction_transfer_poc/2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.xlsx`
- 元Excel: `common/source/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.xlsx`
- 元Excel: `common/source/construction_transfer_poc/2026年02月_作業員勤怠一覧_PoC_時刻確認負例.xlsx`
- 推奨プロファイル: `../config/profiles/v2/construction_transfer_poc.json`
- 特徴:
  - タイトル 3 行 + 4 行目ヘッダー
  - 30 列
  - セル内改行あり
  - 時刻表記ゆれあり
  - 現場名と氏名の微妙な揺れあり
  - `2ページ同一列` と `6ページ同一列` のパターンを同梱
- `ヘッダー不一致負例` で sameHeader 拒否を確認可能
- `時刻確認負例` で `24:30` と片側空の Review を確認可能

## プロファイル雛形生成体験サンプル

- PDF: `common/pdf/profile_wizard_demo/2026年03月_職人別作業日報_Wizard体験.pdf`
- 元Excel: `common/source/profile_wizard_demo/2026年03月_職人別作業日報_Wizard体験.xlsx`
- 利用場所: `samples/v2/pdf/profile_wizard_demo`, `samples/v2/source/profile_wizard_demo`
- 使いどころ:
  - `VER2` の `[6] プロファイル雛形を作成` を体験したいとき
  - 新しい 8 列帳票へ最小手順で対応を始める練習をしたいとき
  - `正規化入場` / `正規化退場` の設定例を学びたいとき

## 再生成方法

次のコマンドで、日本語サンプル帳票一式を再生成できます。

```powershell
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\scripts\build_sample_pdfs.ps1
```

## 試し方

1. 試したい帳票タイプの PDF フォルダを対象にします。
2. 対応するプロファイル JSON を指定して変換します。
3. `Result` シートで日本語の項目名と値が正しく入っていることを確認します。
