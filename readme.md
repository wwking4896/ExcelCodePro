# ExcelCode Pro 使用手冊

## 1. 工具概述

ExcelCode Pro 是一個創新的軟體，可以將 Excel 表格中的數據快速、靈活地轉換為 C/C++ 代碼。本工具支持多檔案、多範圍選擇，並提供強大的程式碼樣板客製化功能。

## 2. 主要功能

- 支持單個或多個 Excel 檔案導入
- 多工作表選擇
- 多資料範圍選擇
- 豐富的程式碼樣板（預設和自訂）
- 程式碼預覽和導出
- 設定檔案保存與載入

## 3. 基本使用流程

### 3.1 選擇 Excel 檔案

1. 點擊「瀏覽...」按鈕選擇單一 Excel 檔案
2. 點擊「多選...」按鈕選擇多個 Excel 檔案

### 3.2 選擇工作表

在下拉選單中選擇要處理的工作表。

### 3.3 選擇資料範圍

1. 點擊「選擇範圍」按鈕
2. 在彈出視窗中輸入 Excel 範圍（如 A1:G10）
3. 點擊「添加範圍」
4. 可添加多個範圍
5. 點擊「確認選擇範圍」

### 3.4 設定程式碼樣板

#### 預設樣板
在下拉選單中選擇預設樣板：
- 陣列初始化
- 二維陣列
- 三維陣列
- 四維陣列（範圍優先/檔案優先）
- 多範圍處理
- 權重表設定

#### 自訂樣板
1. 點擊「設定自訂樣板」
2. 在編輯器中編寫您的樣板
3. 支持多種佔位符

## 4. 樣板佔位符詳解

### 4.1 基本佔位符

| 佔位符 | 說明 | 示例 |
|--------|------|------|
| `{{VALUE}}` | 當前行第一列的值 | `5` |
| `{{ROW_INDEX}}` | 當前行的索引 | `0`, `1`, `2` |
| `{{ALL_COLUMNS}}` | 當前行的所有值 | `1, 2, 3` |

### 4.2 列值佔位符

| 佔位符 | 說明 | 示例 |
|--------|------|------|
| `{{ROW:0}}` | 第一列的值 | `5` |
| `{{ROW:1}}` | 第二列的值 | `10` |
| `{{COL:0}}` | 第一列的值 | `5` |

### 4.3 多範圍佔位符

| 佔位符 | 說明 | 使用場景 |
|--------|------|----------|
| `{{RANGE:1_LOOP_START}}` | 第一個範圍的循環開始 | 多範圍處理 |
| `{{RANGE:1_LOOP_END}}` | 第一個範圍的循環結束 | 多範圍處理 |
| `{{RANGE_1_ROW_COUNT}}` | 第一個範圍的行數 | 範圍大小定義 |

### 4.4 多維度佔位符

| 佔位符 | 說明 | 使用場景 |
|--------|------|----------|
| `{{FILES_LOOP_START}}` | 檔案循環開始 | 多檔案處理 |
| `{{FILES_LOOP_END}}` | 檔案循環結束 | 多檔案處理 |
| `{{RANGES_LOOP_START}}` | 範圍循環開始 | 多範圍處理 |
| `{{RANGES_LOOP_END}}` | 範圍循環結束 | 多範圍處理 |

## 5. 樣板設計範例

### 5.1 一維陣列

```c
unsigned int weights[MAX_SIZE] = {
{{LOOP_START}}
    {{VALUE}},  // 索引 {{ROW_INDEX}}
{{LOOP_END}}
};
```

### 5.2 二維陣列

```c
unsigned int table[ROW_COUNT][COL_COUNT] = {
{{LOOP_START}}
    { {{ALL_COLUMNS}} },
{{LOOP_END}}
};
```

### 5.3 多範圍處理

```c
#define RANGE_1_ROW_COUNT {{RANGE_1_ROW_COUNT}}
#define RANGE_1_COL_COUNT {{RANGE_1_COL_COUNT}}

unsigned int first_area[RANGE_1_ROW_COUNT][RANGE_1_COL_COUNT] = {
{{RANGE:1_LOOP_START}}
    { {{ALL_COLUMNS}} },
{{RANGE:1_LOOP_END}}
};
```

### 5.4 四維陣列（範圍優先）

```c
unsigned int data4d[RANGE_COUNT][FILE_COUNT][ROW_COUNT][COL_COUNT] = {
{{RANGES_LOOP_START}}
    {
{{FILES_LOOP_START}}
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

## 6. 樣板設計技巧

1. 使用巨集定義（如 `MAX_SIZE`, `ROW_COUNT`）
2. 添加註釋提高可讀性
3. 考慮不同維度的數據結構
4. 提供邊界檢查
5. 使用適當的數據類型

## 7. 進階範例：遊戲權重表

```c
void initVariableWeights() {
{{LOOP_START}}
    normal_table_weight[{{COL:0}}][{{COL:1}}][{{COL:2}}][{{COL:3}}] = {{VALUE}};
{{LOOP_END}}
}
```

## 8. 常見應用場景

- 遊戲開發配置
- 科學計算數據
- 多維度實驗數據
- 複雜參數表初始化

## 9. 注意事項

- 確保 Python 3.x（推薦 3.13）
- 需要安裝 pandas、tkinter 等依賴
- 大型 Excel 檔案可能需要較長處理時間

## 10. 技術支持

- GitHub: [專案連結]
- Email: [聯繫方式]

## 結語

Excel 轉程式碼工具提供了一個強大且靈活的解決方案，可以快速將表格數據轉換為結構化的程式碼。通過客製化樣板，您可以輕鬆處理各種複雜的數據轉換需求。

希望本使用手冊能幫助您充分利用這個工具的強大功能！