# ExcelCode Pro

ExcelCode Pro 是一款強大的工具，用於將 Excel 表格資料轉換為 C/C++ 程式碼。透過自訂範本和靈活的範圍選擇，使得從資料到程式碼的轉換過程變得簡單且高效。

## 目錄

1. [功能概述](#功能概述)
2. [程式架構](#程式架構)
3. [安裝與執行](#安裝與執行)
4. [基本使用流程](#基本使用流程)
5. [範本系統](#範本系統)
   - [範本基礎](#範本基礎)
   - [範本標記語法](#範本標記語法)
   - [橫向和直向讀取模式](#橫向和直向讀取模式)
   - [命名範圍](#命名範圍)
   - [多維陣列處理](#多維陣列處理)
6. [範本範例](#範本範例)
7. [進階功能](#進階功能)
8. [疑難排解](#疑難排解)
9. [技術支援](#技術支援)

## 功能概述

ExcelCode Pro 提供以下核心功能：

- **多檔案處理**：可同時處理多個 Excel 檔案
- **工作表選擇**：靈活選擇要處理的工作表
- **範圍管理**：支援多種範圍選擇方式，包括命名範圍
- **靈活的範本系統**：提供多種預設範本及強大的自訂範本功能
- **資料預覽**：即時預覽選定範圍的資料
- **程式碼生成**：自動生成結構化的 C/C++ 程式碼
- **設定檔儲存/載入**：保存和復原完整的工作設定
- **多語言支援**：界面和操作完全支援繁體中文

## 程式架構

ExcelCode Pro 採用模組化設計，主要由以下幾個部分組成：

1. **主程式 (main.py)**：程式入口點，設定基本環境並啟動應用程式
2. **圖形介面 (gui.py)**：處理使用者界面互動和視覺元素展示
3. **Excel 處理器 (excel_handler.py)**：負責 Excel 檔案的載入和資料處理
4. **程式碼生成器 (code_generator.py)**：根據範本和資料生成最終程式碼
5. **公用工具 (utils.py)**：提供各種輔助功能
6. **版本控制 (version.py)**：管理版本和更新檢查

## 安裝與執行

### 系統需求

- Python 3.x (建議 3.13 或更高版本)
- pandas, numpy 等相依套件

### 安裝步驟

1. 下載完整程式碼或執行檔
2. 如果使用原始碼，安裝必要相依套件：
   ```
   pip install -r requirements.txt
   ```

### 執行方式

- 使用執行檔：直接點擊 `ExcelCode Pro.exe`
- 使用原始碼：
  ```
  python main.py
  ```
  或執行 `run.bat`

## 基本使用流程

### 1. 選擇 Excel 檔案

- 點擊「瀏覽...」按鈕選擇單一 Excel 檔案
- 點擊「多選...」按鈕選擇多個 Excel 檔案
- 或從「最近使用的檔案」選單中選取最近使用的檔案

### 2. 選擇工作表

從下拉選單中選擇要處理的工作表。

### 3. 設定範圍

點擊「範圍管理」按鈕，在彈出的對話框中可以：
- 輸入 Excel 範圍格式（如 A1:G10）添加新範圍
- 為範圍設定名稱（方便在範本中引用）
- 預覽所選範圍的實際資料
- 管理多個範圍（添加/刪除/重命名）

### 4. 設定程式碼範本

選擇一種範本方式：
- 從下拉選單中選擇預設範本
- 點擊「設定自訂範本」創建自己的範本
- 點擊「匯入範本」從外部檔案載入範本
- 使用「管理範本」功能管理已儲存的範本

### 5. 生成程式碼

點擊「生成程式碼」按鈕，程式將根據選定的範圍和範本生成 C/C++ 程式碼。

### 6. 儲存結果

- 點擊「儲存程式碼」將生成的程式碼儲存為檔案
- 點擊「複製到剪貼簿」將程式碼複製到剪貼簿

## 範本系統

ExcelCode Pro 的範本系統是其最強大的特性之一，使用者可以通過範本完全控制生成程式碼的結構和格式。

### 範本基礎

範本本質上是一段包含特殊標記（佔位符）的文字，這些標記會在生成程式碼時被實際資料取代。範本可以是任何文字，但通常會是 C/C++ 程式碼架構。

#### 預設範本

系統提供以下預設範本：

1. **陣列初始化**：生成一維陣列
2. **二維陣列**：生成二維陣列（橫向讀取模式）
3. **二維陣列-直向讀取**：生成二維陣列（直向讀取模式）
4. **三維陣列**：適用於多檔案資料的三維陣列
5. **三維陣列-直向讀取**：三維陣列的直向讀取版本
6. **四維陣列 (範圍優先)**：以範圍為主要維度的四維陣列
7. **四維陣列-直向讀取**：四維陣列的直向讀取版本
8. **四維陣列 (檔案優先)**：以檔案為主要維度的四維陣列
9. **三維多範圍陣列**：處理多個範圍的三維陣列
10. **權重表設定**：特化的遊戲權重表格式
11. **權重表設定-直向讀取**：權重表的直向讀取版本
12. **多範圍處理**：處理多個數據範圍的通用模板
13. **命名範圍處理**：使用命名範圍的特化範本

### 範本標記語法

#### 基本標記

| 標記 | 說明 | 示例結果 |
|------|------|----------|
| `{{LOOP_START}}` | 開始資料循環 | - |
| `{{LOOP_END}}` | 結束資料循環 | - |
| `{{VALUE}}` | 當前儲存格值 | `5` |
| `{{ROW_INDEX}}` | 當前行索引 | `0`, `1`, `2` |
| `{{COL_INDEX}}` | 當前列索引 | `0`, `1`, `2` |
| `{{ALL_COLUMNS}}` | 當前行的所有列值 | `1, 2, 3, 4` |
| `{{ALL_ROWS}}` | 當前列的所有行值 | `1, 5, 9, 13` |

#### 方向控制標記

| 標記 | 說明 |
|------|------|
| `{{DIRECTION:ROW}}` | 設定為橫向讀取模式（預設） |
| `{{DIRECTION:COLUMN}}` | 設定為直向讀取模式 |

#### 多範圍標記

| 標記 | 說明 |
|------|------|
| `{{RANGE[範圍名稱]_LOOP_START}}` | 特定命名範圍的循環開始 |
| `{{RANGE[範圍名稱]_LOOP_END}}` | 特定命名範圍的循環結束 |
| `{{RANGE[範圍名稱]_ROW_COUNT}}` | 命名範圍的行數 |
| `{{RANGE[範圍名稱]_COL_COUNT}}` | 命名範圍的列數 |
| `{{RANGE[範圍名稱]_VALUE[行,列]}}` | 命名範圍中特定位置的值 |

#### 檔案相關標記

| 標記 | 說明 |
|------|------|
| `{{FILE_NAME}}` | 當前處理的檔案名稱 |
| `{{FILE_INDEX}}` | 當前檔案索引 |
| `{{FILE_COUNT}}` | 總檔案數量 |

#### 多維陣列特殊標記

| 標記 | 說明 |
|------|------|
| `{{FILES_LOOP_START}}` | 檔案循環開始 |
| `{{FILES_LOOP_END}}` | 檔案循環結束 |
| `{{RANGES_LOOP_START}}` | 範圍循環開始 |
| `{{RANGES_LOOP_END}}` | 範圍循環結束 |
| `{{RANGE_LOOP_START}}` | 範圍內資料循環開始 |
| `{{RANGE_LOOP_END}}` | 範圍內資料循環結束 |

### 橫向和直向讀取模式

ExcelCode Pro 支援兩種資料讀取模式：

#### 橫向讀取模式（預設）

- 資料按行讀取（Row by Row）
- 主要使用 `{{ALL_COLUMNS}}` 獲取一行中的所有值
- 適合表格資料的自然閱讀順序

```
表格資料:
| 1 | 2 | 3 |
| 4 | 5 | 6 |

橫向讀取結果:
行 0: 1, 2, 3
行 1: 4, 5, 6
```

#### 直向讀取模式

- 資料按列讀取（Column by Column）
- 主要使用 `{{ALL_ROWS}}` 獲取一列中的所有值
- 適合需要交換行列結構的情況

```
表格資料:
| 1 | 2 | 3 |
| 4 | 5 | 6 |

直向讀取結果:
列 0: 1, 4
列 1: 2, 5
列 2: 3, 6
```

設定直向讀取模式可通過兩種方式：
1. 在範本中使用 `{{DIRECTION:COLUMN}}` 標記
2. 在範本設定對話框中選擇「直向讀取」選項

### 命名範圍

命名範圍是一種強大的功能，允許為特定範圍指定名稱，然後在範本中引用：

1. 在「範圍管理」中為範圍命名（例如「左上」）
2. 在範本中使用命名範圍標記：
   ```
   #define LEFT_TOP_ROWS {{RANGE[左上]_ROW_COUNT}}
   #define LEFT_TOP_COLS {{RANGE[左上]_COL_COUNT}}

   // 左上角區域數據
   unsigned int left_top_area[LEFT_TOP_ROWS][LEFT_TOP_COLS] = {
   {{RANGE[左上]_LOOP_START}}
       { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
   {{RANGE[左上]_LOOP_END}}
   };
   ```

### 多維陣列處理

處理複雜的多維資料結構時，可利用嵌套循環標記：

```
// 四維陣列: [範圍][檔案][行][列]
unsigned int data4d[RANGE_COUNT][FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{RANGES_LOOP_START}}
    // 範圍 {{RANGE_INDEX}}
    {
{{FILES_LOOP_START}}
        // 來自檔案: {{FILE_NAME}}
        {
{{RANGE_LOOP_START}}
            { {{ALL_COLUMNS}} },
{{RANGE_LOOP_END}}
        },
{{FILES_LOOP_END}}
    },
{{RANGES_LOOP_END}}
};
```

## 範本範例

### 基本一維陣列

```c
// 一維陣列初始化
#define MAX_SIZE 100

unsigned int weights[MAX_SIZE] = {
{{LOOP_START}}
    {{VALUE}},  // 索引 {{ROW_INDEX}}
{{LOOP_END}}
};
```

### 二維陣列（橫向讀取）

```c
// 二維陣列初始化
#define ROW_COUNT 20
#define COL_COUNT 4

unsigned int table[ROW_COUNT][COL_COUNT] = {
{{LOOP_START}}
    { {{ALL_COLUMNS}} },
{{LOOP_END}}
};
```

### 二維陣列（直向讀取）

```c
// 二維陣列初始化 - 直向讀取模式
{{DIRECTION:COLUMN}}  // 指定為直向讀取
#define COL_COUNT 20
#define ROW_COUNT 4

unsigned int table[ROW_COUNT][COL_COUNT] = {
{{LOOP_START}}
    { {{ALL_ROWS}} },  // Column {{COL_INDEX}}
{{LOOP_END}}
};
```

### 命名範圍處理

```c
// 命名範圍資料處理範例
#define LEFT_TOP_ROWS {{RANGE[左上]_ROW_COUNT}}
#define LEFT_TOP_COLS {{RANGE[左上]_COL_COUNT}}
#define RIGHT_TOP_ROWS {{RANGE[右上]_ROW_COUNT}}
#define RIGHT_TOP_COLS {{RANGE[右上]_COL_COUNT}}

// 左上角區域數據
unsigned int left_top_area[LEFT_TOP_ROWS][LEFT_TOP_COLS] = {
{{RANGE[左上]_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE[左上]_LOOP_END}}
};

// 右上角區域數據
unsigned int right_top_area[RIGHT_TOP_ROWS][RIGHT_TOP_COLS] = {
{{RANGE[右上]_LOOP_START}}
    { {{ALL_COLUMNS}} },  // Row {{ROW_INDEX}}
{{RANGE[右上]_LOOP_END}}
};

// 特定欄位值示例
int left_top_first_value = {{RANGE[左上]_VALUE[0,0]}};
int right_top_first_value = {{RANGE[右上]_VALUE[0,0]}};
```

### 權重表設定

```c
// 遊戲權重表初始化
void initVariableWeights() {
{{LOOP_START}}
    normal_table_weight[{{COL:0}}][{{COL:1}}][{{COL:2}}][{{COL:3}}] = {{VALUE}};
{{LOOP_END}}
}
```

### 三維多範圍陣列

```c
// 三維多範圍陣列初始化
#define FILE_COUNT {{FILE_COUNT}}           // 檔案數量
#define RANGE_COUNT {{RANGE_COUNT}}         // 範圍數量
#define MAX_ROW_COUNT {{MAX_ROW_COUNT}}     // 最大行數
#define MAX_COL_COUNT {{MAX_COL_COUNT}}     // 最大列數

// 定義各個範圍的實際大小
unsigned int range_dimensions[RANGE_COUNT][2] = {
{{RANGES_LOOP_START}}
    { {{RANGE_ROW_COUNT}}, {{RANGE_COL_COUNT}} },  // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
{{RANGES_LOOP_END}}
};

// 三維多範圍陣列: [檔案][範圍][行][列]
unsigned int data3d_multi[FILE_COUNT][RANGE_COUNT][MAX_ROW_COUNT][MAX_COL_COUNT] = {
{{FILES_LOOP_START}}
    // 來自檔案: {{FILE_NAME}}
    {
{{RANGES_LOOP_START}}
        // 範圍 {{RANGE_INDEX}}: {{RANGE_STR}}
        {
{{RANGE_DATA_LOOP_START}}
            { {{ALL_COLUMNS}} },
{{RANGE_DATA_LOOP_END}}
        },
{{RANGES_LOOP_END}}
    },
{{FILES_LOOP_END}}
};
```

## 進階功能

### 設定檔儲存與載入

您可以儲存當前的工作設定，包括：
- 已選擇的 Excel 檔案
- 工作表選擇
- 範圍設定（包括命名範圍）
- 範本選擇和內容

使用「儲存設定」和「載入設定」按鈕來管理設定檔。

### 匯入/匯出範本

您可以：
- 將自訂範本儲存為外部檔案
- 從外部檔案載入範本
- 管理多個範本

### 多範圍管理

「範圍管理」工具可以：
- 同時定義和處理多個範圍
- 為範圍命名以便在範本中引用
- 實時預覽範圍內的資料

## 疑難排解

### 常見問題

1. **範本中的標記沒有被正確替換**
   - 檢查標記語法是否正確
   - 確認範圍是否有效且包含資料

2. **生成的程式碼中有逗號問題**
   - 範本中的陣列初始化語法可能需要調整
   - 考慮在範本中明確處理最後一項的逗號

3. **資料類型處理不正確**
   - 預設情況下，程式會嘗試識別數值類型
   - 文字值將被包裹在引號中

### 日誌檢視

程式底部的「執行日誌」區域會顯示重要的操作和錯誤信息，可用於診斷問題。

## 技術支援

如有問題或建議，請聯繫開發者：
- 透過 GitHub 提交 Issue
- 發送電子郵件至 [wu0851202@gmail.com]

---

感謝使用 ExcelCode Pro！希望這個工具能大幅提升您的數據轉程式碼工作效率。
