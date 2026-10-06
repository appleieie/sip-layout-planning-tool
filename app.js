const canvas = document.getElementById('dieCanvas');
const context = canvas.getContext('2d');
const excelInput = document.getElementById('excelInput');
const jsonInput = document.getElementById('jsonInput');
const loadSampleButton = document.getElementById('loadSampleButton');
const exportJsonButton = document.getElementById('exportJsonButton');
const dieNameInput = document.getElementById('dieNameInput');
const dieSizeXInput = document.getElementById('dieSizeXInput');
const dieSizeYInput = document.getElementById('dieSizeYInput');
const padCountValue = document.getElementById('padCountValue');
const rotationValue = document.getElementById('rotationValue');
const zoomValue = document.getElementById('zoomValue');
const sourceValue = document.getElementById('sourceValue');
const statusMessage = document.getElementById('statusMessage');
const zoomInButton = document.getElementById('zoomInButton');
const zoomOutButton = document.getElementById('zoomOutButton');
const fitViewButton = document.getElementById('fitViewButton');

const state = {
  dieAsset: null,
  rotation: 0,
  zoomMultiplier: 1,
  panX: 0,
  panY: 0,
  fontScale: 1,
  sourceName: '-',
};

const GRID_STEP = 500;
const ROTATION_SET = new Set([0, 90, 180, 270]);

function hasXlsxLibrary() {
  return typeof window.XLSX !== 'undefined';
}

function setStatus(message, tone = 'neutral') {
  statusMessage.className = `status-message ${tone}`;
  statusMessage.textContent = message;
}

function stripExtension(fileName) {
  return fileName.replace(/\.[^.]+$/, '');
}

function normalizeDieName(name) {
  if (!name) return 'UnnamedDie';
  const stripped = stripExtension(String(name).trim());
  return stripped.replace(/die$/i, '') || stripped;
}

function round3(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`讀取檔案失敗：${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`讀取檔案失敗：${file.name}`));
    reader.readAsText(file, 'utf-8');
  });
}

function valueToKey(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[_()（）\-]/g, '');
}

function inferDieSizeFromPads(pads) {
  if (!pads.length) {
    return { x: 1000, y: 1000 };
  }

  const maxX = Math.max(...pads.map((pad) => pad.x + pad.width / 2));
  const maxY = Math.max(...pads.map((pad) => pad.y + pad.height / 2));

  return { x: round3(Math.max(maxX, 1)), y: round3(Math.max(maxY, 1)) };
}

function parseNumber(value, fieldName) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    throw new Error(`欄位 ${fieldName} 含有無法解析的數值：${value}`);
  }
  return parsed;
}

function buildDieAsset({ dieName, dieSizeX, dieSizeY, pads }) {
  return {
    format_version: '3.0',
    unit: 'um',
    die: {
      diename: dieName,
      diesize: { x: round3(dieSizeX), y: round3(dieSizeY) },
    },
    pads: pads.map((pad) => ({
      label: String(pad.label ?? ''),
      name: pad.name ?? '',
      x: round3(pad.x),
      y: round3(pad.y),
      width: round3(pad.width),
      height: round3(pad.height),
    })),
  };
}

function parseStandardRows(rows, fallbackName) {
  const headerRowIndex = rows.findIndex((row) => row.some((cell) => valueToKey(cell) === 'diename'));
  if (headerRowIndex === -1) {
    return null;
  }

  const headerRow = rows[headerRowIndex];
  const indexByKey = new Map();
  headerRow.forEach((cell, index) => indexByKey.set(valueToKey(cell), index));

  const required = ['label', 'x', 'y', 'width', 'height'];
  required.forEach((key) => {
    if (!indexByKey.has(key)) {
      throw new Error(`缺少必要欄位：${key}`);
    }
  });

  const pads = [];
  let dieName = '';
  let dieSizeX = null;
  let dieSizeY = null;

  for (const row of rows.slice(headerRowIndex + 1)) {
    if (!row || row.every((cell) => cell == null || String(cell).trim() === '')) {
      continue;
    }

    const label = row[indexByKey.get('label')];
    if (label == null || String(label).trim() === '') {
      continue;
    }

    const currentDieName = row[indexByKey.get('diename')] ?? dieName ?? fallbackName;
    const currentSizeX = row[indexByKey.get('diesizex')];
    const currentSizeY = row[indexByKey.get('diesizey')];

    dieName = String(currentDieName || fallbackName).trim();
    if (currentSizeX != null && currentSizeY != null) {
      dieSizeX = parseNumber(currentSizeX, 'diesize_x');
      dieSizeY = parseNumber(currentSizeY, 'diesize_y');
    }

    pads.push({
      label,
      name: row[indexByKey.get('name')] ?? '',
      x: parseNumber(row[indexByKey.get('x')], 'x'),
      y: parseNumber(row[indexByKey.get('y')], 'y'),
      width: parseNumber(row[indexByKey.get('width')], 'width'),
      height: parseNumber(row[indexByKey.get('height')], 'height'),
    });
  }

  if (!pads.length) {
    throw new Error('Excel 中找不到有效 Pad 資料。');
  }

  const inferredSize = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAsset({
      dieName: dieName || fallbackName,
      dieSizeX: dieSizeX ?? inferredSize.x,
      dieSizeY: dieSizeY ?? inferredSize.y,
      pads,
    }),
    warning: dieSizeX == null || dieSizeY == null ? 'Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。' : '',
  };
}

function parseGroupedSampleRows(rows, fallbackName) {
  const header1 = rows[0] ?? [];
  const signature = [valueToKey(header1[0]), valueToKey(header1[1]), valueToKey(header1[2]), valueToKey(header1[4])].join('|');
  if (signature !== 'padno|name|coordinatexy|padsizexy') {
    return null;
  }

  const pads = rows.slice(2)
    .filter((row) => row && row.some((cell) => cell != null && String(cell).trim() !== ''))
    .map((row) => ({
      label: row[0],
      name: row[1] ?? '',
      x: parseNumber(row[2], 'Coordinate(XY).X'),
      y: parseNumber(row[3], 'Coordinate(XY).Y'),
      width: parseNumber(row[4], 'PadSize(XY).X'),
      height: parseNumber(row[5], 'PadSize(XY).Y'),
    }))
    .filter((pad) => String(pad.label ?? '').trim() !== '');

  if (!pads.length) {
    throw new Error('範例 Excel 中找不到有效 Pad 資料。');
  }

  const inferredSize = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAsset({
      dieName: fallbackName,
      dieSizeX: inferredSize.x,
      dieSizeY: inferredSize.y,
      pads,
    }),
    warning: '此 Excel 未包含 diename / diesize_x / diesize_y，已改用檔名與 Pad 外接範圍推估。',
  };
}

function parseMetaHeaderRows(rows, fallbackName) {
  const META_KEYS = new Set(['diename', 'diesizex', 'diesizey']);

  const meta = {};
  let padHeaderIndex = -1;

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx] ?? [];
    const key0 = valueToKey(row[0]);
    if (META_KEYS.has(key0) || key0 === 'diesize_x' || key0 === 'diesize_y') {
      // scan pairs in this row: col0=key, col1=val, col2=key, col3=val ...
      for (let i = 0; i < row.length - 1; i += 2) {
        const k = valueToKey(row[i]);
        const v = row[i + 1];
        if ((META_KEYS.has(k) || k === 'diesize_x' || k === 'diesize_y') && v != null) {
          meta[k] = v;
        }
      }
    } else if (key0 === 'label') {
      padHeaderIndex = idx;
      break;
    }
  }

  if (padHeaderIndex === -1 || !Object.keys(meta).length) return null;
  if (!('diename' in meta) && !('diesizex' in meta)) return null;

  const dieName = String(meta['diename'] ?? fallbackName).trim() || fallbackName;
  const rawSizeX = meta['diesizex'] ?? meta['diesize_x'];
  const rawSizeY = meta['diesizey'] ?? meta['diesize_y'];
  const dieSizeX = rawSizeX != null ? parseNumber(rawSizeX, 'diesize_x') : null;
  const dieSizeY = rawSizeY != null ? parseNumber(rawSizeY, 'diesize_y') : null;

  const headerRow = rows[padHeaderIndex] ?? [];
  const indexByKey = new Map();
  headerRow.forEach((cell, index) => indexByKey.set(valueToKey(cell), index));

  const required = ['label', 'x', 'y', 'width', 'height'];
  required.forEach((key) => {
    if (!indexByKey.has(key)) throw new Error(`缺少必要 Pad 欄位：${key}`);
  });

  const pads = rows.slice(padHeaderIndex + 1)
    .filter((row) => row && row.some((c) => c != null && String(c).trim() !== ''))
    .map((row) => ({
      label: row[indexByKey.get('label')],
      name: indexByKey.has('name') ? (row[indexByKey.get('name')] ?? '') : '',
      x: parseNumber(row[indexByKey.get('x')], 'x'),
      y: parseNumber(row[indexByKey.get('y')], 'y'),
      width: parseNumber(row[indexByKey.get('width')], 'width'),
      height: parseNumber(row[indexByKey.get('height')], 'height'),
    }))
    .filter((pad) => String(pad.label ?? '').trim() !== '');

  if (!pads.length) throw new Error('meta-header 格式 Excel 中找不到有效 Pad 資料。');

  const inferred = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAsset({
      dieName,
      dieSizeX: dieSizeX ?? inferred.x,
      dieSizeY: dieSizeY ?? inferred.y,
      pads,
    }),
    warning: dieSizeX == null || dieSizeY == null ? 'Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。' : '',
    parser: 'meta-header',
  };
}

function parseWorkbook(buffer, fileName = 'ImportedDie.xlsx') {
  if (!hasXlsxLibrary()) {
    throw new Error('瀏覽器端 XLSX 函式庫未載入。');
  }
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, blankrows: false });
  const fallbackName = normalizeDieName(fileName);
  return (
    parseMetaHeaderRows(rows, fallbackName) ||
    parseStandardRows(rows, fallbackName) ||
    parseGroupedSampleRows(rows, fallbackName) ||
    (() => { throw new Error('無法辨識 Excel 欄位格式。支援格式：\n1. meta-header（前幾列為 diename/diesize，再接 label/x/y 標頭）\n2. 標準格式（單列標頭含 diename 欄）\n3. GBB898die 範例格式（雙列標頭）'); })()
  );
}

async function parseExcelViaServer(buffer, fileName) {
  const response = await fetch(`/api/parse-excel?filename=${encodeURIComponent(fileName)}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/octet-stream',
    },
    body: buffer,
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || '伺服器無法解析 Excel。');
  }

  return payload;
}

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
}

function validateDieAsset(dieAsset) {
  if (!dieAsset || dieAsset.unit !== 'um') {
    throw new Error('JSON 格式錯誤：unit 必須為 um。');
  }
  if (!dieAsset.die?.diename) {
    throw new Error('JSON 格式錯誤：缺少 die.diename。');
  }
  const sizeX = Number(dieAsset.die?.diesize?.x);
  const sizeY = Number(dieAsset.die?.diesize?.y);
  if (!Number.isFinite(sizeX) || sizeX <= 0 || !Number.isFinite(sizeY) || sizeY <= 0) {
    throw new Error('JSON 格式錯誤：die.diesize.x / y 必須大於 0。');
  }
  if (!Array.isArray(dieAsset.pads)) {
    throw new Error('JSON 格式錯誤：pads 必須為陣列。');
  }
}

function syncInputsFromState() {
  if (!state.dieAsset) {
    dieNameInput.value = '';
    dieSizeXInput.value = '';
    dieSizeYInput.value = '';
    return;
  }
  dieNameInput.value = state.dieAsset.die.diename;
  dieSizeXInput.value = state.dieAsset.die.diesize.x;
  dieSizeYInput.value = state.dieAsset.die.diesize.y;
}

function syncStateFromInputs() {
  if (!state.dieAsset) return;
  const name = dieNameInput.value.trim();
  const sizeX = Number(dieSizeXInput.value);
  const sizeY = Number(dieSizeYInput.value);
  if (name) state.dieAsset.die.diename = name;
  if (Number.isFinite(sizeX) && sizeX > 0) state.dieAsset.die.diesize.x = sizeX;
  if (Number.isFinite(sizeY) && sizeY > 0) state.dieAsset.die.diesize.y = sizeY;
}

function transformPoint(x, y, dieSizeX, dieSizeY, rotation) {
  const centerX = dieSizeX / 2;
  const centerY = dieSizeY / 2;
  const dx = x - centerX;
  const dy = y - centerY;
  switch (rotation) {
    case 0: return { x, y };
    case 90: return { x: centerX - dy, y: centerY + dx };
    case 180: return { x: centerX - dx, y: centerY - dy };
    case 270: return { x: centerX + dy, y: centerY - dx };
    default: return { x, y };
  }
}

function getRotatedRectSize(width, height, rotation) {
  return rotation === 90 || rotation === 270 ? { width: height, height: width } : { width, height };
}

function getDieCorners(dieSizeX, dieSizeY, rotation) {
  return [
    transformPoint(0, 0, dieSizeX, dieSizeY, rotation),
    transformPoint(dieSizeX, 0, dieSizeX, dieSizeY, rotation),
    transformPoint(dieSizeX, dieSizeY, dieSizeX, dieSizeY, rotation),
    transformPoint(0, dieSizeY, dieSizeX, dieSizeY, rotation),
  ];
}

function getSceneBounds() {
  if (!state.dieAsset) return null;
  const dieSizeX = Number(state.dieAsset.die.diesize.x);
  const dieSizeY = Number(state.dieAsset.die.diesize.y);
  const dieCorners = getDieCorners(dieSizeX, dieSizeY, state.rotation);
  const points = [...dieCorners];
  state.dieAsset.pads.forEach((pad) => {
    const center = transformPoint(pad.x, pad.y, dieSizeX, dieSizeY, state.rotation);
    const rotatedSize = getRotatedRectSize(pad.width, pad.height, state.rotation);
    points.push(
      { x: center.x - rotatedSize.width / 2, y: center.y - rotatedSize.height / 2 },
      { x: center.x + rotatedSize.width / 2, y: center.y + rotatedSize.height / 2 },
    );
  });
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  return { minX: Math.min(...xs), maxX: Math.max(...xs), minY: Math.min(...ys), maxY: Math.max(...ys) };
}

function resizeCanvasForDisplay() {
  const rect = canvas.getBoundingClientRect();
  const pixelRatio = window.devicePixelRatio || 1;
  const width = Math.max(600, Math.floor(rect.width * pixelRatio));
  const height = Math.max(480, Math.floor(rect.height * pixelRatio));
  if (canvas.width !== width || canvas.height !== height) {
    canvas.width = width;
    canvas.height = height;
  }
}

function drawGrid(bounds, mapX, mapY, scale) {
  const scaledStep = GRID_STEP * scale;
  if (scaledStep < 24) return;
  const startX = Math.floor(bounds.minX / GRID_STEP) * GRID_STEP;
  const endX = Math.ceil(bounds.maxX / GRID_STEP) * GRID_STEP;
  const startY = Math.floor(bounds.minY / GRID_STEP) * GRID_STEP;
  const endY = Math.ceil(bounds.maxY / GRID_STEP) * GRID_STEP;
  context.save();
  context.strokeStyle = '#e5edf6';
  context.lineWidth = 1;
  for (let x = startX; x <= endX; x += GRID_STEP) {
    context.beginPath();
    context.moveTo(mapX(x), mapY(startY));
    context.lineTo(mapX(x), mapY(endY));
    context.stroke();
  }
  for (let y = startY; y <= endY; y += GRID_STEP) {
    context.beginPath();
    context.moveTo(mapX(startX), mapY(y));
    context.lineTo(mapX(endX), mapY(y));
    context.stroke();
  }
  context.restore();
}

function drawAxes(bounds, mapX, mapY, dpr = 1) {
  context.save();
  context.strokeStyle = '#94a3b8';
  context.fillStyle = '#475569';
  context.lineWidth = 1.5;
  const x0 = mapX(0);
  const y0 = mapY(0);
  context.beginPath();
  context.moveTo(x0, mapY(bounds.minY));
  context.lineTo(x0, mapY(bounds.maxY));
  context.stroke();
  context.beginPath();
  context.moveTo(mapX(bounds.minX), y0);
  context.lineTo(mapX(bounds.maxX), y0);
  context.stroke();
  context.font = `${13 * dpr}px Segoe UI`;
  context.fillText('(0,0)', x0 + 8, y0 - 10);
  context.restore();
}

function drawDieOutline(mapX, mapY, dpr = 1) {
  const dieSizeX = Number(state.dieAsset.die.diesize.x);
  const dieSizeY = Number(state.dieAsset.die.diesize.y);
  const corners = getDieCorners(dieSizeX, dieSizeY, state.rotation);
  const screenCorners = corners.map((c) => ({ sx: mapX(c.x), sy: mapY(c.y) }));
  context.save();
  context.beginPath();
  context.moveTo(screenCorners[0].sx, screenCorners[0].sy);
  screenCorners.slice(1).forEach((c) => context.lineTo(c.sx, c.sy));
  context.closePath();
  context.fillStyle = 'rgba(37, 99, 235, 0.06)';
  context.strokeStyle = '#2563eb';
  context.lineWidth = 2;
  context.fill();
  context.stroke();
  // Die Name 標籤：外框上方、左對齊
  const topSy = Math.min(...screenCorners.map((c) => c.sy));
  const leftSx = Math.min(...screenCorners.map((c) => c.sx));
  const fontSize = Math.round(16 * dpr * state.fontScale);
  context.fillStyle = '#1d4ed8';
  context.font = `bold ${fontSize}px Segoe UI`;
  context.textAlign = 'start';
  context.fillText(state.dieAsset.die.diename, leftSx, topSy - 8 * dpr);
  context.restore();
}

function drawPads(mapX, mapY, scale, dpr = 1) {
  const dieSizeX = Number(state.dieAsset.die.diesize.x);
  const dieSizeY = Number(state.dieAsset.die.diesize.y);
  const drawLabels = scale >= 0.12;
  const padFontSize = Math.round(12 * dpr * state.fontScale);
  context.save();
  context.font = `${padFontSize}px Consolas`;
  state.dieAsset.pads.forEach((pad) => {
    const center = transformPoint(pad.x, pad.y, dieSizeX, dieSizeY, state.rotation);
    const rotatedSize = getRotatedRectSize(pad.width, pad.height, state.rotation);
    const left = mapX(center.x - rotatedSize.width / 2);
    const right = mapX(center.x + rotatedSize.width / 2);
    const top = mapY(center.y + rotatedSize.height / 2);
    const bottom = mapY(center.y - rotatedSize.height / 2);
    const rectWidth = right - left;
    const rectHeight = bottom - top;
    context.fillStyle = '#0ea5e9';
    context.strokeStyle = '#0369a1';
    context.lineWidth = 1.2;
    context.fillRect(left, top, rectWidth, rectHeight);
    context.strokeRect(left, top, rectWidth, rectHeight);
    if (drawLabels) {
      context.fillStyle = '#0f172a';
      context.fillText(String(pad.label), left + 4, top + 14 * dpr);
    }
  });
  context.restore();
}

function updateSummary() {
  padCountValue.textContent = state.dieAsset ? String(state.dieAsset.pads.length) : '-';
  rotationValue.textContent = `${state.rotation}°`;
  zoomValue.textContent = `${Math.round(state.zoomMultiplier * 100)}%`;
  sourceValue.textContent = state.sourceName;
}

function centerView() {
  if (!state.dieAsset) { state.panX = 0; state.panY = 0; return; }
  const rect = canvas.getBoundingClientRect();
  const cssW = rect.width || canvas.width / (window.devicePixelRatio || 1);
  const cssH = rect.height || canvas.height / (window.devicePixelRatio || 1);
  const bounds = getSceneBounds();
  const sceneW = Math.max(bounds.maxX - bounds.minX, 1);
  const sceneH = Math.max(bounds.maxY - bounds.minY, 1);
  const PAD = 60;
  const fitScale = Math.min((cssW - PAD * 2) / sceneW, (cssH - PAD * 2) / sceneH);
  const sc = fitScale * state.zoomMultiplier;
  // panX/panY = CSS position of scene (bounds.minX, bounds.minY) on canvas
  state.panX = PAD + ((cssW - PAD * 2) - sceneW * sc) / 2 - bounds.minX * sc;
  state.panY = PAD + ((cssH - PAD * 2) - sceneH * sc) / 2 + bounds.maxY * sc;
}

function drawScene() {
  resizeCanvasForDisplay();
  const dpr = window.devicePixelRatio || 1;
  const width = canvas.width;   // physical px
  const height = canvas.height; // physical px
  const cssW = width / dpr;
  const cssH = height / dpr;

  context.clearRect(0, 0, width, height);
  context.fillStyle = '#f8fbff';
  context.fillRect(0, 0, width, height);

  if (!state.dieAsset) {
    context.fillStyle = '#64748b';
    context.font = `${18 * dpr}px Segoe UI`;
    context.fillText('請先匯入 Excel 或 JSON。', 48, 64);
    updateSummary();
    return;
  }

  const bounds = getSceneBounds();
  const sceneW = Math.max(bounds.maxX - bounds.minX, 1);
  const sceneH = Math.max(bounds.maxY - bounds.minY, 1);
  const PAD = 60;
  const fitScale = Math.min((cssW - PAD * 2) / sceneW, (cssH - PAD * 2) / sceneH);
  const sc = fitScale * state.zoomMultiplier; // CSS px per um

  // mapX/mapY: scene um → physical canvas px
  // state.panX = CSS x position where scene x=0 would be drawn (adjusted for bounds)
  // state.panY = CSS y position where scene y=0 would be drawn (canvas Y increases downward, scene Y increases upward)
  const mapX = (x) => (state.panX + x * sc) * dpr;
  const mapY = (y) => (state.panY - y * sc) * dpr;

  drawGrid(bounds, mapX, mapY, sc * dpr);
  drawAxes(bounds, mapX, mapY, dpr);
  drawDieOutline(mapX, mapY, dpr);
  drawPads(mapX, mapY, sc * dpr, dpr);
  updateSummary();
}

function applyDieAsset(dieAsset, sourceName, warning = '') {
  validateDieAsset(dieAsset);
  state.dieAsset = dieAsset;
  state.sourceName = sourceName;
  state.rotation = 0;
  state.zoomMultiplier = 1;
  syncInputsFromState();
  centerView();
  drawScene();
  setStatus(warning ? `載入完成：${sourceName}\n${warning}` : `載入完成：${sourceName}`, warning ? 'warning' : 'success');
}

async function handleExcelBuffer(buffer, sourceName) {
  try {
    const { dieAsset, warning, parser } = await parseExcelWithFallback(buffer, sourceName);
    applyDieAsset(dieAsset, sourceName, warning);
    if (parser) {
      setStatus(`${sourceName} 載入完成\n解析模式：${parser}${warning ? `\n${warning}` : ''}`, warning ? 'warning' : 'success');
    }
  } catch (error) {
    setStatus(error.message, 'error');
  }
}

async function handleExcelFile(file) {
  const buffer = await readFileAsArrayBuffer(file);
  await handleExcelBuffer(buffer, file.name);
}

async function handleJsonFile(file) {
  try {
    const text = await readFileAsText(file);
    const parsed = JSON.parse(text);
    applyDieAsset(parsed, file.name);
  } catch (error) {
    setStatus(`JSON 載入失敗：${error.message}`, 'error');
  }
}

function exportDieAssetJson() {
  try {
    if (!state.dieAsset) throw new Error('目前沒有可匯出的 Die 資料。');
    syncStateFromInputs();
    validateDieAsset(state.dieAsset);
    const blob = new Blob([JSON.stringify(state.dieAsset, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    const safeName = (state.dieAsset.die.diename || 'die-asset').replace(/[^a-z0-9\-_]+/gi, '_');
    anchor.href = url;
    anchor.download = `${safeName}_die_asset.json`;
    anchor.click();
    URL.revokeObjectURL(url);
    setStatus(`已匯出 JSON：${anchor.download}`, 'success');
  } catch (error) {
    setStatus(`JSON 匯出失敗：${error.message}`, 'error');
  }
}

async function loadBundledSample() {
  try {
    const response = await fetch('/api/sample-die-asset');
    if (!response.ok) {
      throw new Error('無法透過本機伺服器讀取 GBB898die.xlsx。');
    }
    const payload = await response.json();
    if (!payload?.dieAsset) {
      throw new Error(payload.error || '範例 Excel 解析失敗。');
    }
    applyDieAsset(payload.dieAsset, 'GBB898die.xlsx', payload.warning || '');
    setStatus(`GBB898die.xlsx 載入完成\n解析模式：${payload.parser || 'server'}${payload.warning ? `\n${payload.warning}` : ''}`, payload.warning ? 'warning' : 'success');
  } catch (error) {
    setStatus(`載入範例失敗：${error.message}`, 'error');
  }
}

loadSampleButton.addEventListener('click', loadBundledSample);
exportJsonButton.addEventListener('click', exportDieAssetJson);
excelInput.addEventListener('change', async (event) => {
  const [file] = event.target.files ?? [];
  if (file) await handleExcelFile(file);
  event.target.value = '';
});
jsonInput.addEventListener('change', async (event) => {
  const [file] = event.target.files ?? [];
  if (file) await handleJsonFile(file);
  event.target.value = '';
});
[dieNameInput, dieSizeXInput, dieSizeYInput].forEach((input) => {
  input.addEventListener('change', () => {
    syncStateFromInputs();
    drawScene();
  });
});
document.querySelectorAll('[data-rotate]').forEach((button) => {
  button.addEventListener('click', () => {
    const rotation = Number(button.dataset.rotate);
    if (!ROTATION_SET.has(rotation)) return;
    state.rotation = rotation;
    drawScene();
  });
});
zoomInButton.addEventListener('click', () => {
  state.zoomMultiplier = Math.min(state.zoomMultiplier * 1.2, 20);
  drawScene();
});
zoomOutButton.addEventListener('click', () => {
  state.zoomMultiplier = Math.max(state.zoomMultiplier / 1.2, 0.1);
  drawScene();
});
fitViewButton.addEventListener('click', () => {
  state.zoomMultiplier = 1;
  centerView();
  drawScene();
});

const fontScaleValue = document.getElementById('fontScaleValue');
const fontUpButton = document.getElementById('fontUpButton');
const fontDownButton = document.getElementById('fontDownButton');
const fontResetButton = document.getElementById('fontResetButton');

function updateFontScaleLabel() {
  if (fontScaleValue) fontScaleValue.textContent = `${Math.round(state.fontScale * 100)}%`;
}

fontUpButton.addEventListener('click', () => {
  state.fontScale = Math.min(state.fontScale * 1.25, 6);
  updateFontScaleLabel();
  drawScene();
});
fontDownButton.addEventListener('click', () => {
  state.fontScale = Math.max(state.fontScale / 1.25, 0.2);
  updateFontScaleLabel();
  drawScene();
});
fontResetButton.addEventListener('click', () => {
  state.fontScale = 1;
  updateFontScaleLabel();
  drawScene();
});

// ── Pan (drag) support ──────────────────────────────────────
let drag = null;

canvas.addEventListener('mousedown', (e) => {
  if (e.button !== 0) return;
  drag = { startX: e.clientX, startY: e.clientY, panX: state.panX, panY: state.panY };
  canvas.style.cursor = 'grabbing';
  e.preventDefault();
});

window.addEventListener('mousemove', (e) => {
  if (!drag) return;
  state.panX = drag.panX + (e.clientX - drag.startX);
  state.panY = drag.panY + (e.clientY - drag.startY);
  drawScene();
});

window.addEventListener('mouseup', () => {
  if (!drag) return;
  drag = null;
  canvas.style.cursor = 'grab';
});

// Touch pan support
let touchStart = null;

canvas.addEventListener('touchstart', (e) => {
  if (e.touches.length !== 1) return;
  const t = e.touches[0];
  touchStart = { startX: t.clientX, startY: t.clientY, panX: state.panX, panY: state.panY };
  e.preventDefault();
}, { passive: false });

canvas.addEventListener('touchmove', (e) => {
  if (!touchStart || e.touches.length !== 1) return;
  const t = e.touches[0];
  state.panX = touchStart.panX + (t.clientX - touchStart.startX);
  state.panY = touchStart.panY + (t.clientY - touchStart.startY);
  drawScene();
  e.preventDefault();
}, { passive: false });

canvas.addEventListener('touchend', () => { touchStart = null; });

// Mouse-wheel zoom centered on cursor (CSS px)
canvas.addEventListener('wheel', (e) => {
  e.preventDefault();
  const rect = canvas.getBoundingClientRect();
  const cursorCssX = e.clientX - rect.left;
  const cursorCssY = e.clientY - rect.top;
  const factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
  const prevZoom = state.zoomMultiplier;
  const nextZoom = Math.min(Math.max(prevZoom * factor, 0.05), 30);
  // Keep the scene point under cursor fixed
  state.panX = cursorCssX - (cursorCssX - state.panX) * (nextZoom / prevZoom);
  state.panY = cursorCssY - (cursorCssY - state.panY) * (nextZoom / prevZoom);
  state.zoomMultiplier = nextZoom;
  drawScene();
}, { passive: false });
// ─────────────────────────────────────────────────────────────

window.addEventListener('resize', drawScene);
drawScene();
