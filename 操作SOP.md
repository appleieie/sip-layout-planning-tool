# Die 2D Viewer 操作 SOP

## 1. 目的

本文件說明如何啟動與操作 `Die 2D Viewer MVP`，並以 `GBB898die.xlsx` 作為範例，完成 Excel 匯入、2D 顯示、旋轉、縮放、JSON 匯出與 JSON 載入。

## 2. 適用範圍

適用於目前工作目錄中的以下檔案：

- `index.html`
- `styles.css`
- `app.js`
- `serve.py`
- `GBB898die.xlsx`

## 3. 操作前準備

### 3.1 確認檔案位置

請確認以下檔案位於同一資料夾中：

- `index.html`
- `styles.css`
- `app.js`
- `GBB898die.xlsx`
- `serve.py`

### 3.2 確認 Python 可用

本工具的網頁本身不需安裝套件即可執行，但若要使用 `載入範例 Excel` 按鈕，建議先啟動本機 HTTP 伺服器。

## 4. 啟動方式

### 4.1 建議方式：啟動本機伺服器

於目前資料夾執行以下命令：

```python
"d:\部門工作資料\R組工作項目\BP案\2025\凌通A+\晶片Datasheet\SiP Layout planning tool\.venv\Scripts\python.exe" serve.py
```

看到以下訊息代表啟動成功：

```text
Serving ... at http://127.0.0.1:8000
```

接著在瀏覽器開啟：

- `http://127.0.0.1:8000`

### 4.2 備用方式：直接開啟網頁檔

也可直接雙擊 `index.html` 開啟，但這種方式通常無法使用 `載入範例 Excel` 按鈕。若使用此方式，請改用 `匯入 Excel` 手動選檔。

## 5. 基本操作 SOP

### 5.1 載入範例 Excel

1. 開啟網頁後，點選 `載入範例 Excel`。
2. 系統會讀取同資料夾中的 `GBB898die.xlsx`。
3. 成功後，右側畫布會顯示 Die 外框與 Pad 圖形。
4. 左下方 `狀態` 區會顯示載入結果。

### 5.2 手動匯入 Excel

1. 點選 `匯入 Excel`。
2. 選擇 `.xlsx` 或 `.xls` 檔案。
3. 系統會自動判斷格式是否為：
   - v3 標準欄位格式，或
   - `GBB898die.xlsx` 類型的雙列標頭格式。
4. 匯入成功後會自動繪圖。

### 5.3 修改 Die 名稱與尺寸

若 Excel 未提供 `diename`、`diesize_x`、`diesize_y`：

1. 在左側 `Die 資訊` 區手動輸入 `Die Name`。
2. 在 `Die Size X (um)`、`Die Size Y (um)` 輸入正確尺寸。
3. 輸入完成後，畫面會自動更新。

### 5.4 旋轉檢視

1. 在左側 `視圖操作` 區點選：
   - `0°`
   - `90°`
   - `180°`
   - `270°`
2. 系統會以 Die 幾何中心為基準進行旋轉。
3. Pad 與 Die 外框會同步更新方向。

### 5.5 縮放檢視

1. 點選 `放大`：提高檢視倍率。
2. 點選 `縮小`：降低檢視倍率。
3. 點選 `Fit`：重設為適合畫布的檢視比例。

## 6. 匯出 JSON SOP

### 6.1 匯出 `die asset JSON`

1. 確認目前已完成 Excel 或 JSON 載入。
2. 若需要，可先手動修改 `Die Name`、`Die Size X`、`Die Size Y`。
3. 點選 `匯出 die asset JSON`。
4. 瀏覽器會自動下載 JSON 檔案。
5. 預設檔名格式為：

```text
<DieName>_die_asset.json
```

### 6.2 匯出內容

匯出的 JSON 會包含：

- `format_version`
- `unit`
- `die.diename`
- `die.diesize.x`
- `die.diesize.y`
- `pads[]`

## 7. 載入 JSON SOP

1. 點選 `載入 JSON`。
2. 選擇先前匯出的 `die asset JSON` 檔案。
3. 系統會驗證格式。
4. 成功後會重新顯示 Die 圖形。

## 8. 畫面資訊說明

### 8.1 摘要區

左側 `摘要` 會顯示：

- `Pad 數量`
- `目前旋轉`
- `縮放倍率`
- `資料來源`

### 8.2 狀態區

左下方 `狀態` 會顯示：

- 載入成功
- 匯出成功
- 警告訊息
- 錯誤訊息

## 9. 常見問題排除

### 9.1 `載入範例 Excel` 無反應

可能原因：

- 以 `file://` 直接開啟 `index.html`
- 未啟動 `serve.py`

處理方式：

1. 先啟動 `serve.py`
2. 改用 `http://127.0.0.1:8000` 開啟

### 9.2 Excel 匯入失敗

請檢查：

- 是否為 `.xlsx` / `.xls`
- 是否符合支援欄位格式
- 是否有非數字資料出現在座標或尺寸欄位

### 9.3 Die 尺寸看起來不正確

因 `GBB898die.xlsx` 未提供 `diesize_x` / `diesize_y`，系統會先用 Pad 外接範圍推估。若你有正式尺寸，請手動覆寫。

## 10. 建議使用流程

建議實務操作順序如下：

1. 啟動 `serve.py`
2. 開啟 `http://127.0.0.1:8000`
3. 點選 `載入範例 Excel`
4. 確認 Die 外框與 Pad 顯示
5. 必要時修正 `Die Name` 與 `Die Size`
6. 用 `0° / 90° / 180° / 270°` 檢查方向
7. 用 `放大 / 縮小 / Fit` 檢查局部區域
8. 點選 `匯出 die asset JSON`
9. 再用 `載入 JSON` 驗證匯出結果

## 11. 結語

此工具為 MVP 版本，主要目的是快速完成 Die 圖面視覺化與 JSON 資產化。若後續需要加入多 Die 佈局、package layout、碰撞檢查或圖檔輸出，可在下一版持續擴充。
