# V2 サンプル一覧

- 作成日: 2026-03-20 10:45 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20

このフォルダは、`VER2 Secure` の体験と検証に使うサンプル専用の案内です。  
最短で試したい場合は、まず [V2 サンプル体験ガイド](../../docs/02-guides/01-v2-sample-walkthrough.md) を開いてください。

## 1. 使うプロファイル

- 推奨プロファイル: [construction_transfer_poc.json](../../config/profiles/v2/construction_transfer_poc.json)
- メニューで選ぶ場合: `[9] プロファイルを選ぶ` から `construction_transfer_poc`

## 2. サンプルの置き場所

- PDF: [pdf/construction_transfer_poc](pdf/construction_transfer_poc)
- 元 Excel: [source/construction_transfer_poc](source/construction_transfer_poc)
- 雛形生成体験用 PDF: [pdf/profile_wizard_demo](pdf/profile_wizard_demo)
- 雛形生成体験用 元 Excel: [source/profile_wizard_demo](source/profile_wizard_demo)

## 3. サンプル別の見どころ

| サンプル PDF | 用途 | 実感しやすいポイント |
| --- | --- | --- |
| `2026年02月_作業員勤怠一覧_PoC.pdf` | 基本成功 | `Result` に素直に転記される |
| `2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf` | 複数ページ結合 | 同じ列の連続ページをまとめられる |
| `2026年04月-06月_作業員勤怠一覧_PoC_6ページ同一列.pdf` | 長い帳票 | 手結合せずに一括取込できる |
| `2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf` | Review 体験 | 確認が必要な行だけ `Review` に分離される |
| `2026年02月_作業員勤怠一覧_PoC_ヘッダー不一致負例.pdf` | 安全性確認 | 危ない自動結合を拒否できる |
| `2026年03月_職人別作業日報_Wizard体験.pdf` | 雛形生成体験 | 新規帳票向け profile を Wizard で起こせる |

## 4. まず試す順番

1. `2026年02月_作業員勤怠一覧_PoC.pdf`
2. `2026年02月_作業員勤怠一覧_PoC_2ページ同一列.pdf`
3. `2026年02月_作業員勤怠一覧_PoC_時刻確認負例.pdf`
4. `2026年03月_職人別作業日報_Wizard体験.pdf`

## 5. 出力後に見る場所

- `Summary`: PDF 単位の成功・失敗
- `Result`: 取り込めた明細
- `Review`: 人が確認すべき行
- `Errors`: 処理不能だった帳票

## 6. 関連文書

- 体験用手順書: [01-v2-sample-walkthrough.md](../../docs/02-guides/01-v2-sample-walkthrough.md)
- 利用者向け手順: [01-user-manual.md](../../docs/01-current/01-user-manual.md)
