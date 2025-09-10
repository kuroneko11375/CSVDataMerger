# CSV 資料合併器

一個用於將多個 CSV 檔案中特定欄位合併到單一 Excel 檔案的 Python
工具。此工具專門用於處理符合特定命名規則的 CSV
檔案，並擷取預定範圍的資料列。

## 功能特色

-   **圖形化資料夾選擇**：簡單易用的資料夾選擇對話框\
-   **靈活的檔名規則**：可設定的檔案命名樣式\
-   **可自訂資料擷取**：可指定要擷取的欄位與資料列範圍\
-   **錯誤處理**：優雅處理缺失檔案或讀取錯誤\
-   **Excel 輸出**：乾淨的 Excel 檔案輸出，包含正確的欄位名稱\
-   **進度提示**：在終端機輸出處理狀態

## 系統需求

-   Python 3.6+\
-   pandas\
-   tkinter（Python 通常內建）\
-   openpyxl（用於寫入 Excel 檔案）

## 安裝方式

1.  複製專案：

``` bash
git clone https://github.com/yourusername/csv-data-merger.git
cd csv-data-merger
```

2.  安裝必要套件：

``` bash
pip install pandas openpyxl
```

## 使用方法

### 基本用法

1.  執行程式：

``` bash
python csv_merger.py
```

2.  在彈出的對話框中選擇包含 CSV 檔案的資料夾

3.  程式會處理符合命名規則的檔案，並輸出合併後的 Excel 檔案

### 檔名規則

預設程式會尋找以下樣式的檔案：

    DATA1_RESULT_YYYYMMDD_*.csv
    DATA2_RESULT_YYYYMMDD_*.csv
    DATA3_RESULT_YYYYMMDD_*.csv
    DATA4_RESULT_YYYYMMDD_*.csv
    DATA5_RESULT_YYYYMMDD_*.csv
    DATA6_RESULT_YYYYMMDD_*.csv

其中 `YYYYMMDD` 為今日日期，例如：`20250910`。

### 自訂參數

你可以在 `main()` 函式中修改參數以改變程式行為：

``` python
merge_csv_files(
    folder_path=folder_selected,
    file_prefix="YOUR_PREFIX",  # 修改檔案前綴字
    column_index=3,             # 要擷取的欄位 (0=A, 1=B, 2=C, 3=D, ...)
    start_row=55,               # 起始列 (1-based，55 = 第 55 列)
    num_rows=12                 # 要擷取的列數
)
```

### 參數說明

-   **file_prefix**：CSV 檔名前綴字（預設：`DATA`）\
-   **column_index**：要擷取的欄位（0 為 A 欄，3 為 D 欄，以此類推）\
-   **start_row**：起始列號（以 1 為基準，55 代表第 55 列）\
-   **num_rows**：從起始列開始要擷取的資料列數

## 輸出結果

程式會在選擇的資料夾中建立一個名為 `merged_data_YYYYMMDD.xlsx` 的 Excel
檔案，內容包含：

-   每個處理過的檔案各一欄（DATA1, DATA2, ...）\
-   每個檔案指定欄位的擷取資料列\
-   若檔案缺失或讀取失敗，則建立空欄位

## 範例

假設你有以下檔案：

    SENSOR1_RESULT_20250910_morning.csv
    SENSOR2_RESULT_20250910_morning.csv
    SENSOR3_RESULT_20250910_morning.csv
    SENSOR4_RESULT_20250910_morning.csv
    SENSOR5_RESULT_20250910_morning.csv
    SENSOR6_RESULT_20250910_morning.csv

可在程式中修改：

``` python
file_prefix="SENSOR"
```

## 錯誤處理

程式會處理多種錯誤情況： - 缺失檔案：建立空欄位\
- 檔案讀取錯誤：輸出錯誤訊息並建立空欄位\
- 無效資料：繼續處理其他檔案

## 貢獻方式

1.  Fork 專案\
2.  建立新分支 (`git checkout -b feature/amazing-feature`)\
3.  提交變更 (`git commit -m 'Add amazing feature'`)\
4.  推送分支 (`git push origin feature/amazing-feature`)\
5.  建立 Pull Request

## 授權條款

本專案使用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案。

## 更新紀錄

### 版本 1.0.0

-   初始版本\
-   基本 CSV 合併功能\
-   圖形化資料夾選擇\
-   可自訂參數\
-   錯誤處理與提示

## 支援

若遇到問題或有疑問，請在 GitHub 上建立 issue。

## 作者

Your Name - your.email@example.com

專案連結：<https://github.com/yourusername/csv-data-merger>
