# Die 晶片腳位 2D 繪圖與 SiP 佈局規劃程式開發說明書

本文件旨在定義與實作一隻用於讀取 Excel 檔案中 Die（晶片）腳位資訊，並支援 **晶片資料存取 (Save/Load)** 與 **SiP (System in Package) 多晶片放置設計** 的工具程式。

---

## 1. 專案概述

[cite_start]在半導體封裝設計與測試階段，晶片腳位（Pad）的實體位置與尺寸資訊通常會以 Excel 表格形式記錄 [cite: 3][cite_start]。為了方便工程師進行快速視覺化檢查、確認走線可行性以及防呆驗證，我們需要開發一隻自動化工具，能夠一鍵將結構化的 Excel 數據轉換為精準的 2D 幾何圖形 [cite: 3]。

### 核心功能目標
* [cite_start]**精準繪圖**：嚴格依據 微米（$\mu m$）單位進行繪圖，避免比例失真 [cite: 3]。
* **晶片資產化 (JSON)**：支援將 Excel 讀入的晶片資料儲存為標準 JSON 檔，方便後續直接載入，無須重複讀取 Excel。
* **SiP 多晶片並存 (Package Canvas)**：支援在同一個主畫布（Substrate）上載入多顆不同的晶片，並可為每顆晶片指定獨立的 **平移 offset (X, Y)** 與 **旋轉角度 (Rotation)**。

---

## 2. 資料格式規格

### 2.1 原始 Excel 欄位 (單位：$\mu m$)
[cite_start]程式將讀取包含以下欄位的 Excel 工作表 [cite: 3]：

| 欄位名稱 | 資料型態 | 說明 | 範例資料 |
| :--- | :--- | :--- | :--- |
| **標號** | 整數 / 字串 | [cite_start]腳位的序號或唯一代碼，如 `1`, `2`, `P1` [cite: 3] | [cite_start]`1` [cite: 3] |
| **名稱** | 字串 | [cite_start]腳位的訊號名稱或功能定義 [cite: 3] | [cite_start]`VDD_CORE` [cite: 3] |
| **X座標** | 浮點數 | [cite_start]Pad 的中心點在 Die 上的 X 軸座標 [cite: 3] | [cite_start]`500.0` [cite: 3] |
| **Y座標** | 浮點數 | [cite_start]Pad 的中心點在 Die 上的 Y 軸座標 [cite: 3] | [cite_start]`1200.5` [cite: 3] |
| **Pad尺寸(X軸)** | 浮點數 | [cite_start]Pad 本身在 X 軸方向的寬度 (Width) [cite: 4] | [cite_start]`60.0` [cite: 4] |
| **Pad尺寸(Y軸)** | 浮點數 | [cite_start]Pad 本身在 Y 軸方向的高度 (Height) [cite: 4] | [cite_start]`60.0` [cite: 4] |

### 2.2 晶片儲存格式 (JSON)
單一晶片儲存時，會轉換為如下的結構化 JSON 檔案：
```json
{
  "die_name": "Controller_Die",
  "pads": [
    {
      "label": "1",
      "name": "VDD_CORE",
      "x": 500.0,
      "y": 1200.5,
      "width": 60.0,
      "height": 60.0
    }
  ]
}