# Die 2D Viewer MVP

這是一個純靜態 `HTML + CSS + JavaScript` 的 MVP 工具，用來：

- 匯入 Excel Die / Pad 資料
- 顯示 Die 2D 版圖
- 支援 `0° / 90° / 180° / 270°` 旋轉
- 支援等比例縮放
- 匯出與載入 `die asset JSON`

## 檔案

- `index.html`：主畫面
- `styles.css`：樣式
- `app.js`：資料解析、繪圖、旋轉、縮放、JSON 匯出/載入
- `serve.py`：本機靜態伺服器
- `validate_sample.py`：用 `GBB898die.xlsx` 產生 `sample_die_asset.json` 的驗證腳本
- `GBB898die.xlsx`：範例 Excel

## 啟動方式

建議使用本機 HTTP 伺服器，這樣 `載入範例 Excel` 按鈕才能直接抓到 `GBB898die.xlsx`。

```python
"d:\部門工作資料\R組工作項目\BP案\2025\凌通A+\晶片Datasheet\SiP Layout planning tool\.venv\Scripts\python.exe" serve.py
```

開啟瀏覽器後進入：

- `http://127.0.0.1:8000`

## 使用方式

1. 點選 `載入範例 Excel`，直接讀取 `GBB898die.xlsx`
2. 或點選 `匯入 Excel` 選擇自有檔案
3. 若 Excel 未包含 `diename` 與 `diesize`，可在左側欄位手動調整
4. 使用 `0° / 90° / 180° / 270°` 切換方向
5. 使用 `放大 / 縮小 / Fit` 調整視圖
6. 點選 `匯出 die asset JSON` 下載結果
7. 點選 `載入 JSON` 回讀先前匯出的檔案

## 範例驗證

以下命令會解析 `GBB898die.xlsx`，並輸出 `sample_die_asset.json`：

```python
"d:\部門工作資料\R組工作項目\BP案\2025\凌通A+\晶片Datasheet\SiP Layout planning tool\.venv\Scripts\python.exe" validate_sample.py
```

## 注意事項

- `GBB898die.xlsx` 採用雙列標頭格式：
  - 第 1 列：`PadNo`, `Name`, `Coordinate(XY)`, `PadSize(XY)`
  - 第 2 列：`X`, `Y`, `X`, `Y`
- 此範例檔未內含 `diename`、`diesize_x`、`diesize_y`，工具會自動改用檔名與 Pad 外接範圍推估。
- 若直接以 `file://` 開啟 `index.html`，`載入範例 Excel` 按鈕可能無法讀取本機檔案；此時請改用 `匯入 Excel` 或啟動 `serve.py`。
