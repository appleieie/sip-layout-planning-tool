"""
下載 xlsx.full.min.js 並打包成 standalone.html
執行方式：雙擊本檔案，或在命令列執行 python download_xlsx_and_build.py
"""
import re
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).parent

XLSX_URL = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"
XLSX_PATH = ROOT / "xlsx.full.min.js"


def download_xlsx():
    if XLSX_PATH.exists() and XLSX_PATH.stat().st_size > 100_000:
        print(f"xlsx.full.min.js 已存在 ({XLSX_PATH.stat().st_size // 1024} KB)，跳過下載。")
        return
    print("正在下載 xlsx.full.min.js ...")
    urllib.request.urlretrieve(XLSX_URL, XLSX_PATH)
    print(f"下載完成：{XLSX_PATH.stat().st_size // 1024} KB")


def build():
    css_text   = (ROOT / "styles.css").read_text(encoding="utf-8")
    xlsx_js    = XLSX_PATH.read_text(encoding="utf-8")
    app_js_src = (ROOT / "app.js").read_text(encoding="utf-8")

    # ── 移除 parseExcelViaServer ──────────────────────────────
    app_js_src = re.sub(
        r"async function parseExcelViaServer\(.*?\n\}\n",
        "",
        app_js_src,
        flags=re.DOTALL,
    )

    # ── parseExcelWithFallback → 純前端 ───────────────────────
    OLD_FALLBACK = """\
async function parseExcelWithFallback(buffer, fileName) {
  try {
    return await parseExcelViaServer(buffer, fileName);
  } catch (serverError) {
    if (hasXlsxLibrary()) {
      try {
        return parseWorkbook(buffer, fileName);
      } catch (clientError) {
        throw new Error(`Excel 解析失敗。伺服器：${serverError.message}；瀏覽器：${clientError.message}`);
      }
    }
    throw serverError;
  }
}"""
    NEW_FALLBACK = """\
async function parseExcelWithFallback(buffer, fileName) {
  // Standalone: client-side only
  return parseWorkbook(buffer, fileName);
}"""
    app_js_src = app_js_src.replace(OLD_FALLBACK, NEW_FALLBACK)

    # ── loadBundledSample → 提示 ──────────────────────────────
    app_js_src = re.sub(
        r"async function loadBundledSample\(\) \{.*?\n\}",
        """\
async function loadBundledSample() {
  setStatus('Standalone 版本不支援此功能。\\n請使用「匯入 Excel」按鈕選擇本機檔案。', 'warning');
}""",
        app_js_src,
        flags=re.DOTALL,
    )

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Die 2D Viewer MVP</title>
  <style>
{css_text}
  </style>
</head>
<body>
  <div class="app-shell">
    <aside class="sidebar">
      <div>
        <h1>Die 2D Viewer</h1>
        <p class="subtitle">MVP：Excel 匯入、2D 顯示、直角旋轉、等比例縮放、JSON 匯出/載入</p>
      </div>

      <section class="panel">
        <h2>資料來源</h2>
        <div class="button-group">
          <button id="loadSampleButton" type="button" class="secondary">載入範例 Excel</button>
          <label class="file-button">
            匯入 Excel
            <input id="excelInput" type="file" accept=".xlsx,.xls" hidden>
          </label>
          <label class="file-button secondary">
            載入 JSON
            <input id="jsonInput" type="file" accept=".json" hidden>
          </label>
          <button id="exportJsonButton" type="button" class="secondary">匯出 die asset JSON</button>
        </div>
      </section>

      <section class="panel">
        <h2>Die 資訊</h2>
        <div class="field-grid">
          <label>
            Die Name
            <input id="dieNameInput" type="text" placeholder="例如：GBB898">
          </label>
          <label>
            Die Size X (um)
            <input id="dieSizeXInput" type="number" min="0" step="0.001" placeholder="例如：5200">
          </label>
          <label>
            Die Size Y (um)
            <input id="dieSizeYInput" type="number" min="0" step="0.001" placeholder="例如：4000">
          </label>
        </div>
        <p class="hint">若 Excel 未提供 diename、diesize_x、diesize_y，系統會以 Pad 外接範圍推估，可在此手動修正。</p>
      </section>

      <section class="panel">
        <h2>視圖操作</h2>
        <div class="button-group compact">
          <button type="button" data-rotate="0">0°</button>
          <button type="button" data-rotate="90">90°</button>
          <button type="button" data-rotate="180">180°</button>
          <button type="button" data-rotate="270">270°</button>
        </div>
        <div class="button-group compact top-gap">
          <button id="zoomInButton" type="button">放大</button>
          <button id="zoomOutButton" type="button">縮小</button>
          <button id="fitViewButton" type="button" class="secondary">Fit</button>
        </div>
        <div class="button-group compact top-gap">
          <button id="fontUpButton" type="button">字A+</button>
          <button id="fontDownButton" type="button">字A-</button>
          <button id="fontResetButton" type="button" class="secondary">字重設</button>
          <span id="fontScaleValue" class="badge">100%</span>
        </div>
      </section>

      <section class="panel">
        <h2>摘要</h2>
        <dl class="summary-list">
          <div><dt>Pad 數量</dt><dd id="padCountValue">-</dd></div>
          <div><dt>目前旋轉</dt><dd id="rotationValue">0°</dd></div>
          <div><dt>縮放倍率</dt><dd id="zoomValue">100%</dd></div>
          <div><dt>資料來源</dt><dd id="sourceValue">-</dd></div>
        </dl>
      </section>

      <section class="panel status-panel">
        <h2>狀態</h2>
        <div id="statusMessage" class="status-message neutral">尚未載入資料。</div>
      </section>
    </aside>

    <main class="canvas-area">
      <div class="canvas-toolbar">
        <span>座標原點：左下角</span>
        <span>單位：um</span>
        <span>繪圖：等比例縮放</span>
      </div>
      <canvas id="dieCanvas" width="1200" height="860"></canvas>
    </main>
  </div>

  <script>
/* ── XLSX.js (bundled) ── */
{xlsx_js}
  </script>
  <script>
/* ── Die 2D Viewer (standalone) ── */
{app_js_src}
  </script>
</body>
</html>
"""

    out = ROOT / "standalone.html"
    out.write_text(html, encoding="utf-8")
    size_kb = out.stat().st_size // 1024
    print(f"✅  standalone.html 產生完成！")
    print(f"    路徑：{out}")
    print(f"    大小：{size_kb} KB")
    print()
    print("使用方式：")
    print("  直接雙擊 standalone.html 用 Chrome / Edge 開啟。")
    print("  複製這一個 .html 檔案到任何電腦即可使用，不需要安裝任何東西。")


if __name__ == "__main__":
    try:
        download_xlsx()
        build()
    except Exception as e:
        print(f"❌ 發生錯誤：{e}", file=sys.stderr)
        input("按 Enter 關閉...")
        sys.exit(1)
    input("按 Enter 關閉...")
