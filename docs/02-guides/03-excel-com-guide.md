# Excel COM オートメーション 学習ガイド

- 作成日: 2026-03-15 23:49 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-20
- 目的: Excel COM の基礎知識と PDF2Excel での採用リスクを学ぶ

---

## 1. COM とは何か

### 1-1. 定義

**COM (Component Object Model)** は、Microsoft が 1993 年に策定した、アプリケーション間通信の標準規格です。

```
┌────────────────────┐     COM プロトコル       ┌────────────────────┐
│  クライアント       │ ◀───────────────────▶   │  サーバー           │
│  (PowerShell)      │   プロセス間通信 (IPC)   │  (Excel.exe)       │
│                    │   インターフェース呼出    │                    │
└────────────────────┘                         └────────────────────┘
```

| 用語 | 意味 |
|---|---|
| COM クライアント | COM オブジェクトを**使う**側（今回は PowerShell） |
| COM サーバー | COM オブジェクトを**提供する**側（今回は Excel） |
| COM オブジェクト | 操作対象のインスタンス（Workbook, Worksheet, Range 等） |
| IPC | Inter-Process Communication。プロセス間でデータや命令をやり取りする仕組み |
| CLSID | クラス ID。COM オブジェクトをレジストリから見つけるための GUID |
| ProgID | プログラム ID。`Excel.Application` のような人間が読める名前 |

### 1-2. なぜ COM が生まれたか

| 時代 | 課題 | COM の解決策 |
|---|---|---|
| 1990 年代 | Word の文書に Excel のグラフを埋め込みたい | OLE（Object Linking and Embedding）として COM を基盤に実現 |
| 2000 年代 | スクリプトから Office を操作して業務自動化したい | VBScript / VBA / PowerShell から COM 経由で操作 |
| 2010 年代〜 | Web やクラウドから Office を使いたい | → Microsoft Graph API / Office JS へ移行が進む（COM は Windows ローカル専用） |

### 1-3. COM の仕組み（簡易版）

```
① PowerShell が「Excel.Application を起動したい」とOSに要求
     ↓
② OS がレジストリで HKCR\Excel.Application の CLSID を検索
     ↓
③ 対応する CLSID の InprocServer32 / LocalServer32 を確認
     ↓
④ Excel.exe を起動（Out-of-Process COM サーバー）
     ↓
⑤ RPC (Remote Procedure Call) チャネルを確立
     ↓
⑥ PowerShell から Excel のメソッドを呼び出し可能になる
```

> **Out-of-Process vs In-Process**
>
> | 種類 | 説明 | 例 |
> |---|---|---|
> | Out-of-Process | 別プロセスとして起動 | Excel.Application（Excel.exe が起動する） |
> | In-Process | 呼び出し元のプロセス内で動作 | Shell32.dll 等の DLL |
>
> Excel COM は **Out-of-Process** です。つまり PowerShell と Excel は別プロセスとして動き、プロセス間通信で命令を送受信します。

---

## 2. PowerShell から Excel COM を使う

### 2-1. 基本パターン

```powershell
# ① Excel プロセスを起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false          # 画面に表示しない
$excel.DisplayAlerts = $false    # 確認ダイアログを出さない

# ② ブックを操作
$workbook = $excel.Workbooks.Open("C:\test.xlsx")
$worksheet = $workbook.Worksheets.Item("Sheet1")

# ③ セルを読み書き
$value = $worksheet.Range("A1").Value2     # 読み取り
$worksheet.Range("B1").Value2 = "Hello"    # 書き込み

# ④ 保存して閉じる
$workbook.Save()
$workbook.Close($false)
$excel.Quit()

# ⑤ COM オブジェクトを解放（最重要！）
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
```

### 2-2. COM オブジェクトの階層構造

```
Excel.Application
  ├── Workbooks (コレクション)
  │     ├── Workbook (1冊のブック)
  │     │     ├── Worksheets (コレクション)
  │     │     │     ├── Worksheet (1枚のシート)
  │     │     │     │     ├── Range (セル範囲)
  │     │     │     │     │     ├── Value2 (値)
  │     │     │     │     │     ├── Text (表示テキスト)
  │     │     │     │     │     ├── Font (フォント設定)
  │     │     │     │     │     └── Interior (背景色)
  │     │     │     │     ├── Cells (全セル)
  │     │     │     │     ├── UsedRange (使用範囲)
  │     │     │     │     ├── ListObjects (テーブル)
  │     │     │     │     ├── Columns (列)
  │     │     │     │     └── PageSetup (印刷設定)
  │     │     │     └── ...
  │     │     ├── Queries (Power Query)
  │     │     │     ├── Add() (クエリ追加)
  │     │     │     └── Item() (クエリ取得)
  │     │     ├── VBProject (VBA プロジェクト)
  │     │     │     └── VBComponents (VBA モジュール)
  │     │     └── RefreshAll() (全クエリ更新)
  │     └── ...
  ├── Run() (VBA マクロ実行)
  └── Quit() (Excel 終了)
```

### 2-3. よく使うプロパティとメソッド

| オブジェクト | プロパティ/メソッド | 説明 | PDF2Excel での使用場面 |
|---|---|---|---|
| `Application` | `.Visible` | Excel ウィンドウの表示/非表示 | `$false` で裏で動作させる |
| `Application` | `.DisplayAlerts` | 確認ダイアログの表示/非表示 | `$false` で無人実行 |
| `Application` | `.AutomationSecurity` | マクロのセキュリティ | `1` = 有効にする |
| `Application` | `.Quit()` | Excel を終了 | 最終クリーンアップ |
| `Application` | `.Run(macro)` | VBA マクロを実行 | `ExportResultAsXlsx` の呼び出し |
| `Workbooks` | `.Open(path)` | ブックを開く | runtime xlsm を開く |
| `Workbooks` | `.Add()` | 新規ブックを作成 | テンプレート生成時 |
| `Workbook` | `.Queries.Add(name, formula)` | Power Query を追加 | M 式を注入 |
| `Workbook` | `.SaveAs(path, format)` | 名前を付けて保存 | 形式 51=xlsx, 52=xlsm |
| `Workbook` | `.RefreshAll()` | 全クエリを更新 | VBA 側で呼び出し |
| `Workbook` | `.ExportAsFixedFormat(0, path)` | PDF として出力 | テスト用 PDF 生成 |
| `Workbook` | `.Close(save)` | ブックを閉じる | `$false` = 保存せず閉じる |
| `Worksheet` | `.Range("A1")` | セル範囲を取得 | セルの読み書き |
| `Worksheet` | `.Cells.Item(row, col)` | 行列指定でセル取得 | ループ内のセルアクセス |
| `Worksheet` | `.UsedRange` | 使用済みセル範囲 | 行数・列数の計測 |
| `Worksheet` | `.ListObjects` | テーブル一覧 | tblResult 等のテーブル操作 |
| `Worksheet` | `.Columns.AutoFit()` | 列幅を自動調整 | 見栄えの調整 |
| `Range` | `.Value2` | セルの値（日付は数値） | 読み書きの基本 |
| `Range` | `.Text` | 表示されるテキスト | 書式込みの表示値 |
| `Range` | `.NumberFormat` | 表示形式 | `'@'` = テキスト形式 |
| `Range` | `.Font.Bold` | 太字 | ヘッダー行の装飾 |
| `Range` | `.Interior.Color` | 背景色 | `15773696` = 薄青 |
| `Range` | `.EntireColumn` | 列全体 | 列単位の書式設定 |
| `ListObject` | `.DataBodyRange.Rows.Count` | データ行数 | Result/Errors の行数計測 |
| `ListObject` | `.QueryTable` | テーブルの QueryTable | Power Query の結果をシートに展開 |
| `QueryTable` | `.Refresh($false)` | クエリを同期更新 | `$false` = 完了まで待つ |

### 2-4. Value2 と Value の違い

| プロパティ | 日付セル | 通貨セル | 速度 |
|---|---|---|---|
| `.Value` | `DateTime` 型で返る | `Currency` 型で返る | やや遅い |
| `.Value2` | 浮動小数点数（シリアル値）で返る | `Double` 型で返る | **速い** |

PDF2Excel では一貫して `.Value2` を使用しています。日付の解釈が不要で、パフォーマンスが優先されるためです。

---

## 3. COM の参照カウントとメモリ管理

### 3-1. 参照カウント方式

COM オブジェクトは **参照カウント** でメモリ管理されます。

```
       参照追加          参照追加
PowerShell ────▶ Worksheet ────▶ Range("A1")
  refcount=1       refcount=1      refcount=1

       解放              解放
PowerShell ────▶ Worksheet ────▶ Range("A1")
  refcount=0       refcount=0      refcount=0
  → 解放           → 解放          → 解放
```

| 操作 | 参照カウント |
|---|---|
| 変数に代入 | +1 |
| `FinalReleaseComObject()` | → 0 に強制（即解放） |
| `ReleaseComObject()` | -1（0 になったら解放） |
| 変数がスコープ外 | GC 任せ（いつ解放されるか不定） |

### 3-2. PDF2Excel の解放パターン

PDF2Excel では `Release-ComObject` パイプライン関数を用意しています。

```powershell
# pdf2excel.common.ps1 で定義
function Release-ComObject {
    param([Parameter(ValueFromPipeline = $true)]$InputObject)
    process {
        if ($null -ne $InputObject -and
            [System.Runtime.InteropServices.Marshal]::IsComObject($InputObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($InputObject)
        }
    }
}
```

使い方:

```powershell
# パイプラインで送ると null チェック + IsComObject チェック付きで安全に解放
$worksheet | Release-ComObject
$workbook  | Release-ComObject
$excel     | Release-ComObject
```

### 3-3. Two-Dot Rule（二点ルール）

COM プログラミングの有名な罠です。

```powershell
# ❌ 悪い例: ドットが 2 つ以上連なる
$value = $workbook.Worksheets.Item("Sheet1").Range("A1").Value2
# → Worksheets コレクション、Worksheet、Range の 3 つが暗黙に作られ、
#   変数に入っていないため解放できない

# ✅ 良い例: 各オブジェクトを変数に受ける
$worksheets = $workbook.Worksheets
$worksheet  = $worksheets.Item("Sheet1")
$range      = $worksheet.Range("A1")
$value      = $range.Value2
# → 全て変数に入っているため、確実に解放できる
$range      | Release-ComObject
$worksheet  | Release-ComObject
$worksheets | Release-ComObject
```

> **PDF2Excel での対応**:
> 実用上のバランスとして、PDF2Excel では長い `try/finally` ブロック内で作業し、`finally` で一括解放 + `[GC]::Collect()` を呼ぶパターンを採用しています。完全な Two-Dot Rule の遵守より、コードの読みやすさとの両立を優先しています。

---

## 4. PDF2Excel での Excel COM の使われ方

### 4-1. COM を使う場面（全 4 箇所）

| # | ファイル | 目的 | COM の生存時間 |
|---|---|---|---|
| 1 | `build_excel_template.ps1` | テンプレート xlsm の生成 | 数秒（シート作成 → VBA 取込 → 保存） |
| 2 | `run_pdf2excel.ps1` (メイン) | PDF 変換の本体処理 | **数分**（ブック操作 → PQ 更新 → 後処理 → 出力） |
| 3 | `build_sample_pdfs.ps1` | テスト用 PDF の生成 | 数秒（Excel → PDF エクスポート） |
| 4 | `run_integration_tests.ps1` | テスト用 fixture 生成 + 結果読取 | 数秒〜数十秒 |

### 4-2. メイン処理での COM ライフサイクル

```
Open-ExcelRuntimeContext          ← ① Excel 起動 + ブック読込
    ↓
Configure-WorkbookQueries         ← ② Power Query を注入
    ↓
Load-WorkbookOutputSheets         ← ③ QueryTable.Refresh で PDF 読込（最も時間がかかる）
    ↓
Apply-VersionSpecificWorkbookEnrichments  ← ④ セルを 1 行ずつ読み書き
    ↓
Measure-WorkbookOutcome           ← ⑤ 各シートの行数を計測
    ↓
Update-WorkbookSummaryState       ← ⑥ Control/Summary に結果を書込
    ↓
Publish-WorkbookOutput            ← ⑦ VBA マクロで xlsx 出力
    ↓
Close-ExcelRuntimeContext         ← ⑧ ブック閉じ → Quit → COM 解放
```

### 4-3. COM オブジェクトの総数（メイン処理 1 回あたりの概算）

| オブジェクト種別 | 概数 | 生成元 |
|---|---|---|
| `Application` | 1 | `New-Object -ComObject` |
| `Workbook` | 1 | `Workbooks.Open()` |
| `Worksheet` | 5〜10 | `Worksheets.Item()` の呼び出し回数 |
| `Range` | 100〜500 | セルの読み書き回数（正規化処理のループ） |
| `ListObject` | 5〜7 | テーブルの作成・確認 |
| `QueryTable` | 5〜6 | Power Query の展開 |
| `VBProject` 系 | 2〜3 | テンプレート生成時のみ |
| **合計** | **約 120〜530** | 正規化対象行数に比例 |

---

## 5. PDF2Excel で Excel COM を採用するリスク

### 5-1. リスク一覧

| # | リスクカテゴリ | リスク内容 | 深刻度 | 発生頻度 | PDF2Excel での対策 |
|---|---|---|---|---|---|
| R-1 | **プロセスリーク** | COM 解放漏れで Excel.exe がゾンビとして残る | 🔴 高 | 中 | `try/finally` + `Release-ComObject` + `[GC]::Collect()` + テストでプロセス残留チェック |
| R-2 | **排他エラー** | 別のスクリプトや手動 Excel がファイルをロック | 🟡 中 | 高 | 名前付き Mutex で同時実行を防止。ただし手動 Excel は防げない |
| R-3 | **ダイアログブロック** | 予期せぬ確認ダイアログで処理が停止 | 🔴 高 | 低 | `DisplayAlerts = $false` と `AskToUpdateLinks = $false` で抑制 |
| R-4 | **バージョン依存** | Excel のバージョンで API の挙動が微妙に異なる | 🟡 中 | 低 | M365 に限定。`Pdf.Tables` は Excel 2019+ / M365 が必要 |
| R-5 | **セキュリティポリシー** | 企業環境でマクロ実行が禁止されている | 🟠 中高 | 中 | VBA マクロ失敗時の PowerShell フォールバック |
| R-6 | **パフォーマンス** | COM のプロセス間通信は遅い（特にセル単位のループ） | 🟡 中 | 常時 | 正規化処理で顕著。数百行×定義数のセルアクセスが発生 |
| R-7 | **環境依存** | Windows + デスクトップ版 Excel でのみ動作 | 🔴 高 | 該当時 | Linux/Mac/Web版Excel では動作不可能。設計上の制約として許容 |
| R-8 | **クラッシュリカバリ** | 処理中に Excel がクラッシュすると一時ファイルが残る | 🟡 中 | 低 | 次回起動時に `Compact-RuntimeArtifacts` で掃除 |
| R-9 | **テスト困難性** | COM 操作のモックが困難でユニットテストしにくい | 🟡 中 | 常時 | 統合テストで実際の Excel を起動して検証 |
| R-10 | **同時セッション** | ターミナルサービス（RDP等）で複数ユーザーが実行 | 🟡 中 | 低 | `Global\` Mutex でセッション横断の排他 |

### 5-2. 各リスクの詳細

---

#### R-1. プロセスリーク（最も重要）

**問題**: COM オブジェクトを正しく解放しないと、Excel.exe がバックグラウンドに残り続ける。

```
正常時:
   PowerShell → Excel.exe 起動 → 処理 → Quit() → Excel.exe 終了

異常時:
   PowerShell → Excel.exe 起動 → 例外発生 → Quit() されず → Excel.exe が残留
                                                                    ↑
                                                           タスクマネージャーに
                                                           EXCEL.exe が表示される
```

**影響**:
- メモリを消費し続ける（1 プロセスあたり 50-200MB）
- ファイルのロックが残る
- 次回の実行が正常に動かない

**PDF2Excel の対策**:

```powershell
# try/finally パターン（run_pdf2excel.ps1 末尾）
try {
    # ... 本処理 ...
} catch {
    # エラー時もControl シートに '失敗' を書いて保存を試みる
} finally {
    Close-ExcelRuntimeContext -RuntimeContext $runtimeContext  # ← 確実に実行
    Finalize-RunWorkspace
    Release-RunLock
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
```

```powershell
# Close-ExcelRuntimeContext の中身
function Close-ExcelRuntimeContext {
    param($RuntimeContext)
    if ($null -eq $RuntimeContext) { return }

    # Workbook を閉じる（エラーは無視）
    if ($RuntimeContext.Workbook) {
        try { $RuntimeContext.Workbook.Close($false) } catch {}
    }
    # Excel を終了（エラーは無視）
    if ($RuntimeContext.Excel) {
        try { $RuntimeContext.Excel.Quit() } catch {}
    }
    # COM オブジェクトを個別解放（エラーは無視）
    foreach ($comObject in @($RuntimeContext.SummarySheet,
                              $RuntimeContext.ControlSheet,
                              $RuntimeContext.Workbook,
                              $RuntimeContext.Excel)) {
        try { $comObject | Release-ComObject } catch {}
    }
}
```

**テストでの検証**:

```powershell
# 統合テストの最後に Excel プロセス残留チェック
'Excel プロセス残留なし' {
    $leaked = @(Wait-For-ExcelBaseline -BaselineIds $suiteBaselineExcel -TimeoutSeconds 10)
    Assert-True -Condition ($leaked.Count -eq 0)
        -Message ("Excel process leak detected: " + ($leaked -join ', '))
}
```

---

#### R-2. 排他エラー

**問題**: Excel ファイルは排他ロックで開かれるため、同じファイルを複数プロセスが同時に開けない。

**PDF2Excel の対策**:

| 対策 | 説明 |
|---|---|
| テンプレートのコピー | テンプレート原本ではなく、`runtime/` にコピーした一時ファイルを操作する |
| Mutex | 同時実行そのものを防止する（OS レベルの排他） |
| run.lock ファイル | ロック中の情報（誰が、いつから）を可視化する |

---

#### R-6. パフォーマンス問題

**問題**: COM のプロセス間通信は 1 回あたりマイクロ秒単位のオーバーヘッドがある。セル単位のループでは顕著。

```powershell
# 正規化処理のループ（Add-V2ResultNormalizedColumns 内）
for ($row = 2; $row -le $rowCount; $row += 1) {
    foreach ($definition in $definitions) {
        # ← この1行で COM 往復が 1 回発生
        $rawValue = [string]$Worksheet.Cells.Item($row, $headerMap[$definition.SourceColumnName]).Text
        # ← この1行でも COM 往復が 1 回発生
        $Worksheet.Cells.Item($row, $headerMap[$definition.DisplayName]).Value2 = $normalized.NormalizedText
    }
}
# 15 行 × 4 定義 × 読み書き 3 回 = 約 180 回の COM 往復
```

**高速化の理論（現在は未採用）**:

| 方法 | 概要 | 効果 |
|---|---|---|
| `Range.Value2` の一括読取 | `$data = $range.Value2` で 2 次元配列として一括取得 | 読取が N 回 → 1 回 |
| `Range.Value2` の一括書込 | 2 次元配列をまとめてアサイン | 書込が N 回 → 1 回 |
| クリップボード経由 | TSV 文字列をコピペ | 大量データに有効 |

一括操作の例:

```powershell
# ❌ 遅い: セルを1つずつ読む
for ($i = 1; $i -le 100; $i++) {
    $values += $worksheet.Cells.Item($i, 1).Value2    # COM往復 100回
}

# ✅ 速い: 範囲を一括で読む
$range = $worksheet.Range("A1:A100")
$array = $range.Value2     # COM往復 1回。$array は [100, 1] の2次元配列
```

---

#### R-7. 環境依存

**問題**: Excel COM は以下の条件をすべて満たす環境でのみ動作します。

| 条件 | 必要な環境 | 満たさない場合 |
|---|---|---|
| OS | Windows | Linux / macOS では COM 自体が存在しない |
| Excel | デスクトップ版 Excel | Web 版 Excel では COM を公開していない |
| Excel バージョン | 2019 以降 / M365 | `Pdf.Tables()` が使えない古い Excel |
| 権限 | COM オートメーション許可 | サーバー環境や一部の企業ポリシーで禁止 |
| VBA | マクロ実行許可（推奨） | VBA なしでも PowerShell フォールバックで動作 |

---

## 6. Excel COM の代替技術と比較

| 技術 | PDF 読取 | Excel 書込 | Windows 必須 | Excel 必須 | 速度 |
|---|---|---|---|---|---|
| **Excel COM（現行）** | ✅ Pdf.Tables() | ✅ 完全互換 | ✅ 必須 | ✅ 必須 | 🟡 中 |
| EPPlus / ClosedXML | ❌ PDF 不可 | ✅ xlsx 生成可 | ❌ 不要 | ❌ 不要 | 🟢 速い |
| Python (openpyxl) | ❌ PDF 不可 | ✅ xlsx 生成可 | ❌ 不要 | ❌ 不要 | 🟢 速い |
| Python (tabula-py) | ⚠️ 別エンジン | ⚠️ 別ライブラリ | ❌ 不要 | ❌ 不要 | 🟡 中 |
| Office JS (Web Add-in) | ❌ PDF 不可 | ✅ Web Excel | ❌ 不要 | ⚠️ Web版 | 🟡 中 |
| Microsoft Graph API | ❌ PDF 不可 | ⚠️ 限定的 | ❌ 不要 | ❌ 不要 | 🔴 遅い |

> **なぜ COM を採用しているか？**
>
> PDF2Excel が COM を使う最大の理由は **`Pdf.Tables()`** です。この関数は Excel の Power Query エンジンにのみ搭載されており、COM 経由でしかプログラムから利用できません。代替の PDF 読取ライブラリ（tabula-py 等）は抽出精度や対応フォーマットが異なり、Excel の `Pdf.Tables()` と同じ結果を保証できません。

---

## 7. デバッグのコツ

### 7-1. Excel プロセスの確認

```powershell
# 実行中の Excel プロセスを一覧表示
Get-Process EXCEL -ErrorAction SilentlyContinue | Format-Table Id, StartTime, MainWindowTitle

# ゾンビ Excel を強制終了（注意: 他の Excel も終了する）
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

### 7-2. COM オブジェクトの種類を確認

```powershell
$excel = New-Object -ComObject Excel.Application

# 型名を確認
$excel.GetType().FullName
# → System.__ComObject

# COM オブジェクトかどうかを確認
[System.Runtime.InteropServices.Marshal]::IsComObject($excel)
# → True
```

### 7-3. Excel を表示して動作確認

```powershell
$excel.Visible = $true   # デバッグ時に表示すると何が起きてるか見える
```

### 7-4. エラーハンドリング

```powershell
try {
    $worksheet.Range("ZZ99999").Value2 = "test"
} catch [System.Runtime.InteropServices.COMException] {
    Write-Host "COM エラー: $($_.Exception.Message)"
    Write-Host "HRESULT: 0x$($_.Exception.ErrorCode.ToString('X8'))"
}
```

| HRESULT | 意味 |
|---|---|
| `0x800A03EC` | 範囲外のセルアクセス |
| `0x800A01A8` | オブジェクトがこのプロパティをサポートしていない |
| `0x80010105` | サーバー（Excel）との RPC 接続が切れた |
| `0x800AC472` | ファイルが別プロセスでロックされている |

---

## 8. まとめ

### COM のメリット

| メリット | 説明 |
|---|---|
| **Pdf.Tables() が使える** | COM 経由でしかアクセスできない Excel 独自のPDF解析機能 |
| **完全な Excel 互換** | マクロ、Power Query、書式、テーブル等の全機能にアクセス可能 |
| **追加インストール不要** | Windows + Excel があれば動作する |
| **VBA との相互運用** | PowerShell から VBA を呼び、VBA から PowerShell の結果を使える |

### COM のデメリット

| デメリット | 説明 |
|---|---|
| **Windows + Excel 専用** | クロスプラットフォームでは使えない |
| **プロセスリークの危険性** | 解放漏れで Excel.exe が残る |
| **パフォーマンスの限界** | プロセス間通信のオーバーヘッド |
| **テストの困難さ** | モック化できず、統合テストが必須 |
| **ダイアログの罠** | 想定外のポップアップで処理が止まる |
| **将来性の不安** | Microsoft は Graph API / Office JS 推しに移行中 |

### PDF2Excel の設計判断

```
PDF2Excel が COM を使い続ける理由:
  ┌─ Pdf.Tables() に代替がない
  ├─ 対象ユーザーは Windows + Excel M365 を持っている
  ├─ リスクは全て対策済み（Mutex, try/finally, フォールバック等）
  └─ 業務ツールとしてはこのアプローチが最も実用的
```
