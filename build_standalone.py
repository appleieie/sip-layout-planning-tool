"""
build_standalone.py
將 styles.css、xlsx.full.min.js、app.js 與 index.html 打包成
單一 standalone.html，可在任何電腦上直接雙擊開啟，無需 Python 伺服器。
"""
from pathlib import Path
import re

ROOT = Path(__file__).parent

# ── 讀取來源 ───────────────────────────────────────────────────
css_text    = (ROOT / "styles.css").read_text(encoding="utf-8")
xlsx_js     = (ROOT / "xlsx.full.min.js").read_text(encoding="utf-8")
app_js_src  = (ROOT / "app.js").read_text(encoding="utf-8")

# ── 修改 app.js：移除伺服器 API 相依 ──────────────────────────
# 1. 移除 parseExcelViaServer 函式
app_js_src = re.sub(
    r"async function parseExcelViaServer\(.*?\n\}\n",
    "",
    app_js_src,
    flags=re.DOTALL,
)

# 2. 將 parseExcelWithFallback 改為純前端解析
app_js_src = app_js_src.replace(
    """\
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
}""",
    """\
async function parseExcelWithFallback(buffer, fileName) {
  // Standalone mode: client-side only
  return parseWorkbook(buffer, fileName);
}"""
)

# 3. 將 loadBundledSample 改為提示使用者
app_js_src = re.sub(
    r"async function loadBundledSample\(\) \{.*?\n\}",
    """\
async function loadBundledSample() {
  setStatus('Standalone 版本不支援「載入範例 Excel」。\\n請使用「匯入 Excel」按鈕選擇本機檔案。', 'warning');
}""",
    app_js_src,
    flags=re.DOTALL,
)

# ── 組裝 HTML ──────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Die 2D Viewer MVP (Standalone)</title>
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

  <!-- XLSX.js (bundled offline) -->
  <script>
{xlsx_js}
  </script>

  <!-- App logic (standalone, no server API) -->
  <script>
// ── Stub: loadSampleButton is removed in standalone; guard against missing element
const loadSampleButton = {{ addEventListener: () => {{}} }};
{app_js_src}
  </script>
</body>
</html>
"""

out = ROOT / "standalone.html"
out.write_text(html, encoding="utf-8")
print(f"✅  standalone.html 已產生：{out}")
print(f"    大小：{out.stat().st_size / 1024:.0f} KB")
print()
print("使用方式：")
print("  直接雙擊 standalone.html 用瀏覽器開啟即可。")
print("  不需要 Python、不需要網路、不需要任何安裝。")
print("  只要把這一個 .html 檔案複製到任何電腦。")
