# SEO 自動排名追蹤工具

自動化 Google 搜尋排名追蹤工具，用於監測**特定目標網域**在各關鍵字搜尋結果中的排名位置，並以 Pivot 矩陣格式輸出至 Excel。

## 功能特點

- **目標網域追蹤**：追蹤指定網域在 Google 搜尋結果中的排名
- **多頁搜尋**：可設定搜尋深度（預設 2 頁），跨頁累計排名
- **Pivot 矩陣輸出**：日期為列、關鍵字為欄，交叉處顯示排名（如 `6/P.1`）
- **定價管理**：支援雙階定價（tier1/tier2），含稅自動計算
- **假日自動判定**：週末及國定假日自動標記為「休」
- **專業 Excel 格式**：凍結窗格、條件填充、框線、置中對齊
- **防遺失儲存**：每個關鍵字搜尋完成即儲存
- **無痕模式搜尋**：使用 Chrome 無痕模式，減少個人化搜尋影響

## 系統需求

- **作業系統**：macOS（包含 AppleScript 整合）
- **Python 版本**：Python 3.7+
- **瀏覽器**：Google Chrome
- **ChromeDriver**：與 Chrome 版本相容的 ChromeDriver

## 安裝步驟

### 1. 建立虛擬環境

```bash
cd seo_auto_check
python3 -m venv venv
source venv/bin/activate
```

### 2. 安裝相依套件

```bash
./venv/bin/pip install openpyxl selenium
```

### 3. 確認 Chrome 已安裝

確保系統已安裝 Google Chrome 瀏覽器。

## 使用方法

### 步驟 1：設定 CONFIG 工作表

開啟 `data/input/seo_search_keyword.xlsx`，切換至 **CONFIG** 工作表：

| 欄位 | 值 | 說明 |
|------|-----|------|
| `target_domain` | `newgaotec.com` | 要追蹤的目標網域 |
| `search_depth_pages` | `2` | 搜尋幾頁 Google 結果 |
| `outside_threshold` | `15` | 超過此排名顯示為「N以外」 |
| `holidays` | `2026-01-01, ...` | 國定假日清單（逗號分隔） |
| `output_filename` | `seo_search_results.xlsx` | 輸出檔名 |

**重要**：`target_domain` 必須填寫，否則程式無法執行。

### 步驟 2：設定關鍵字與定價

切換至 **KEYWORDS** 工作表：

| A 欄（編號） | B 欄：KEYWORDS | C 欄：tier1_condition | D 欄：tier1_price | E 欄：tier2_condition | F 欄：tier2_price |
|-------------|----------------|----------------------|-------------------|----------------------|-------------------|
| 1           | 紅外光譜儀      | 10                   | 1000              |                      |                   |
| 2           | 分光光度計      | 10                   | 1000              |                      |                   |
| 3           | 恆溫恆濕箱      | 10                   | 2900              | 15                   | 1900              |

- **tier1_condition / tier1_price**：第一階條件名次與未稅價格（必填）
- **tier2_condition / tier2_price**：第二階條件名次與未稅價格（選填）

### 步驟 3：執行程式

```bash
./venv/bin/python seo_auto_check.py
```

### 步驟 4：查看結果

開啟 `data/output/seo_search_results.xlsx` 查看結果。

## 檔案結構

```
seo_auto_check/
├── seo_auto_check.py              # 主程式
├── data/
│   ├── input/
│   │   └── seo_search_keyword.xlsx  # 設定 + 關鍵字（CONFIG / KEYWORDS 工作表）
│   └── output/
│       └── seo_search_results.xlsx  # 搜尋結果輸出
├── venv/                           # Python 虛擬環境
└── README.md
```

### 輸出 Excel 格式

工作表以 `YYYYMM` 命名（如 `202601`），結構如下：

```
     A              B            C          ...  R            S
1  條件(名次)      [tier2]      [tier2]     ...  [tier2]      **未超過11個關鍵字不收費**
2  價格(未稅)      [tier2價格]                   [tier2價格]
3  價格(含稅)      =B2*1.05     ...              =R2*1.05     [青色填充]
4  條件(名次)      10           10          ...  10
5  價格(未稅)      [tier1價格]  [tier1價格] ...  [tier1價格]
6  價格(含稅)      =B5*1.05     =C5*1.05    ...  =R5*1.05     [青色填充]
7  關鍵字          紅外光譜儀    分光光度計   ...  FTIR ATR
8  01-01-26       休            休           ...  休
9  01-02-26       6/P.1        2/P.1        ...  2/P.1        11:39
...
```

- 排名格式：`6/P.1`（第 6 名，第 1 頁）
- 超出範圍：`15以外`
- 假日標記：`休`

## 執行範例

```
============================================================
  SEO 自動排名追蹤工具
  目標網域排名監測 — Pivot 矩陣輸出
============================================================

目標網域: newgaotec.com
搜尋深度: 2 頁
超出閾值: 15
關鍵字數: 17

[1/17] 搜尋關鍵字：紅外光譜儀
  ✓ 找到！排名第 6 名，第 1 頁
  結果: 6/P.1

[2/17] 搜尋關鍵字：分光光度計
  ✓ 找到！排名第 2 名，第 1 頁
  結果: 2/P.1
...

============================================================
✓ 所有關鍵字搜尋完成！共處理 17 個關鍵字
✓ 結果已儲存至 data/output/seo_search_results.xlsx
============================================================
```

## 注意事項

### 使用限制

1. **搜尋頻率**：關鍵字之間設有 5 秒等待時間，避免觸發 Google 反機器人機制
2. **每日一次**：同一天重複執行會被跳過，避免覆蓋資料
3. **作業系統**：目前僅支援 macOS（使用 AppleScript 關閉瀏覽器）
4. **搜尋結果**：因 Google 搜尋結果會受地區、時間等因素影響，結果可能有所不同

### 建議事項

- **定期執行**：建議每日固定時間執行，以獲得一致的比較基準
- **備份資料**：定期備份 `seo_search_results.xlsx`，避免資料遺失
- **關鍵字管理**：妥善維護 `seo_search_keyword.xlsx`，保持關鍵字列表更新

## 技術架構

### 核心套件

- **Selenium**：瀏覽器自動化控制
- **openpyxl**：Excel 檔案讀寫
- **Chrome WebDriver**：Chrome 瀏覽器驅動

### 主要功能模組

| 函式 | 說明 |
|------|------|
| `load_config()` | 從 CONFIG 工作表讀取設定 |
| `load_keywords()` | 讀取關鍵字 + 定價資料 |
| `is_holiday()` | 判斷日期是否為假日 |
| `format_ranking()` | 格式化排名為 `N/P.X` 字串 |
| `search_keyword()` | 執行 Google 搜尋，尋找目標網域排名 |
| `init_month_sheet()` | 建立月份工作表（含標題列 + 預填日期） |
| `save_ranking()` | 寫入排名資料到對應儲存格 |
| `format_sheet()` | 套用 Excel 格式（框線、填充、對齊） |
| `main()` | 主程式流程控制 |

## 故障排除

### 找不到 ChromeDriver

確保系統已安裝 ChromeDriver 或由 Selenium 自動管理。

### CONFIG 未設定 target_domain

程式啟動時會檢查，若未填寫會提示錯誤訊息。

### Excel 檔案無法開啟

確保執行程式時沒有同時開啟 `seo_search_results.xlsx`。

### 今天已有資料

同一天重複執行會跳過，如需重新執行請先手動清除當日資料。

## 授權

本專案僅供內部使用，請勿用於商業用途或大量自動化搜尋。
