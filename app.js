// ============================================================
// Die / SiP 2D Viewer — app.js
// 支援：單晶片檢視（legacy）＋ 多晶片 SiP 佈局（拖拉對齊、打線繪圖）
// ============================================================

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
const sipPanel = document.getElementById('sipPanel');
const activeDieSelect = document.getElementById('activeDieSelect');
const die2OffsetXInput = document.getElementById('die2OffsetXInput');
const die2OffsetYInput = document.getElementById('die2OffsetYInput');
const sipSizeXInput = document.getElementById('sipSizeXInput');
const sipSizeYInput = document.getElementById('sipSizeYInput');
const dieListSummary = document.getElementById('dieListSummary');

const GRID_STEP = 500;
const ROTATION_SET = new Set([0, 90, 180, 270]);
const RAIL_GAP = 3000; // 兩條外框線間隔 3mm
const RAIL_INNER_MARGIN = 1000; // 晶片群外緣到內圈線的距離
const SIP_AUTO_MARGIN = 600; // SiP 外框預設在外圈線外再留的邊界

const state = {
  mode: 'single', // 'single' | 'sip'
  dies: [], // { sheet, diename, diesize:{x,y}, pads:[...], rotation, posX, posY }
  activeDieIndex: 0,
  sip: { sizeX: 0, sizeY: 0, posX: 0, posY: 0, auto: true },
  zoomMultiplier: 1,
  panX: 0,
  panY: 0,
  fontScale: 1,
  tooltipFontScale: 1.4, // 資訊方塊字體倍率
  sourceName: '-',
  selectedPad: null, // { dieIndex, label } 點擊選取的腳位
};

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
  if (!pads.length) return { x: 1000, y: 1000 };
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

function buildDieAssetRaw({ dieName, dieSizeX, dieSizeY, pads }) {
  return {
    diename: dieName,
    diesize: { x: round3(dieSizeX), y: round3(dieSizeY) },
    pads: pads.map((pad) => {
      const item = {
        label: String(pad.label ?? ''),
        name: pad.name ?? '',
        x: round3(pad.x),
        y: round3(pad.y),
        width: round3(pad.width),
        height: round3(pad.height),
      };
      if (pad.wirebond != null && String(pad.wirebond).trim() !== '') {
        item.wirebond = String(pad.wirebond).trim();
      }
      return item;
    }),
  };
}

// ── Client-side Excel 解析（伺服器不可用時 fallback） ──────────

function parseMetaHeaderRows(rows, fallbackName) {
  const META_KEYS = new Set(['diename', 'diesizex', 'diesizey']);
  const meta = {};
  let padHeaderIndex = -1;

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx] ?? [];
    const key0 = valueToKey(row[0]);
    if (META_KEYS.has(key0) || key0 === 'diesize_x' || key0 === 'diesize_y') {
      for (let i = 0; i < row.length - 1; i += 2) {
        const k = valueToKey(row[i]);
        const v = row[i + 1];
        if ((META_KEYS.has(k) || k === 'diesize_x' || k === 'diesize_y') && v != null) meta[k] = v;
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
      wirebond: indexByKey.has('wirebond') ? row[indexByKey.get('wirebond')] : undefined,
    }))
    .filter((pad) => String(pad.label ?? '').trim() !== '');

  if (!pads.length) throw new Error('meta-header 格式 Excel 中找不到有效 Pad 資料。');

  const inferred = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAssetRaw({
      dieName,
      dieSizeX: dieSizeX ?? inferred.x,
      dieSizeY: dieSizeY ?? inferred.y,
      pads,
    }),
    warning: dieSizeX == null || dieSizeY == null ? 'Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。' : '',
    parser: 'meta-header',
  };
}

function parseStandardRows(rows, fallbackName) {
  const headerRowIndex = rows.findIndex((row) => row.some((cell) => valueToKey(cell) === 'diename'));
  if (headerRowIndex === -1) return null;

  const headerRow = rows[headerRowIndex];
  const indexByKey = new Map();
  headerRow.forEach((cell, index) => indexByKey.set(valueToKey(cell), index));

  const required = ['label', 'x', 'y', 'width', 'height'];
  required.forEach((key) => {
    if (!indexByKey.has(key)) throw new Error(`缺少必要欄位：${key}`);
  });

  const pads = [];
  let dieName = '';
  let dieSizeX = null;
  let dieSizeY = null;

  for (const row of rows.slice(headerRowIndex + 1)) {
    if (!row || row.every((cell) => cell == null || String(cell).trim() === '')) continue;
    const label = row[indexByKey.get('label')];
    if (label == null || String(label).trim() === '') continue;
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
      wirebond: indexByKey.has('wirebond') ? row[indexByKey.get('wirebond')] : undefined,
    });
  }

  if (!pads.length) throw new Error('Excel 中找不到有效 Pad 資料。');

  const inferredSize = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAssetRaw({
      dieName: dieName || fallbackName,
      dieSizeX: dieSizeX ?? inferredSize.x,
      dieSizeY: dieSizeY ?? inferredSize.y,
      pads,
    }),
    warning: dieSizeX == null || dieSizeY == null ? 'Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。' : '',
    parser: 'standard',
  };
}

function parseGroupedSampleRows(rows, fallbackName) {
  const header1 = rows[0] ?? [];
  const signature = [valueToKey(header1[0]), valueToKey(header1[1]), valueToKey(header1[2]), valueToKey(header1[4])].join('|');
  if (signature !== 'padno|name|coordinatexy|padsizexy') return null;

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

  if (!pads.length) throw new Error('範例 Excel 中找不到有效 Pad 資料。');

  const inferredSize = inferDieSizeFromPads(pads);
  return {
    dieAsset: buildDieAssetRaw({
      dieName: fallbackName,
      dieSizeX: inferredSize.x,
      dieSizeY: inferredSize.y,
      pads,
    }),
    warning: '此 Excel 未包含 diename / diesize_x / diesize_y，已改用檔名與 Pad 外接範圍推估。',
    parser: 'grouped-sample',
  };
}

function parseSheetRows(rows, fallbackName) {
  return (
    parseMetaHeaderRows(rows, fallbackName) ||
    parseStandardRows(rows, fallbackName) ||
    parseGroupedSampleRows(rows, fallbackName)
  );
}

function parseWorkbookClient(buffer, fileName) {
  if (!hasXlsxLibrary()) throw new Error('瀏覽器端 XLSX 函式庫未載入。');
  const workbook = XLSX.read(buffer, { type: 'array' });
  const fallbackName = normalizeDieName(fileName);

  const parsedSheets = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, blankrows: false });
    let result = null;
    try {
      result = parseSheetRows(rows, sheetName);
    } catch (err) {
      result = null;
    }
    if (result) parsedSheets.push({ sheetName, result });
  }

  if (parsedSheets.length >= 2) {
    const dies = parsedSheets.map(({ sheetName, result }) => ({
      sheet: sheetName,
      diename: result.dieAsset.diename,
      diesize: result.dieAsset.diesize,
      pads: result.dieAsset.pads,
    }));
    const warning = parsedSheets
      .filter(({ result }) => result.warning)
      .map(({ sheetName, result }) => `${sheetName}: ${result.warning}`)
      .join('\n');
    return { sipAsset: { format_version: '4.0', unit: 'um', dies }, warning, parser: 'multi-sheet-client' };
  }

  if (parsedSheets.length === 1) {
    const { result } = parsedSheets[0];
    if (result.dieAsset.diename === workbook.SheetNames[0]) result.dieAsset.diename = fallbackName;
    return result;
  }

  throw new Error(
    '無法辨識 Excel 欄位格式。\n支援格式：\n1. meta-header（前幾列為 diename/diesize，再接 label/x/y/width/height 標頭）\n2. 標準格式（單列標頭含 diename 欄）\n3. GBB898die 範例格式（雙列標頭）\n多工作表檔案：每個 sheet 各自符合上述格式即可組成 SiP。',
  );
}

async function parseExcelViaServer(buffer, fileName) {
  const response = await fetch(`/api/parse-excel?filename=${encodeURIComponent(fileName)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: buffer,
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error || '伺服器無法解析 Excel。');
  return payload;
}

async function parseExcelWithFallback(buffer, fileName) {
  try {
    return await parseExcelViaServer(buffer, fileName);
  } catch (serverError) {
    if (hasXlsxLibrary()) {
      try {
        return parseWorkbookClient(buffer, fileName);
      } catch (clientError) {
        throw new Error(`Excel 解析失敗。伺服器：${serverError.message}；瀏覽器：${clientError.message}`);
      }
    }
    throw serverError;
  }
}

// ── 幾何／座標工具 ──────────────────────────────────────────

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

function getDieWorldRect(die) {
  const cx = die.posX + die.diesize.x / 2;
  const cy = die.posY + die.diesize.y / 2;
  const rotatedSize = getRotatedRectSize(die.diesize.x, die.diesize.y, die.rotation);
  return {
    minX: cx - rotatedSize.width / 2,
    maxX: cx + rotatedSize.width / 2,
    minY: cy - rotatedSize.height / 2,
    maxY: cy + rotatedSize.height / 2,
    cx,
    cy,
  };
}

function getPadWorldCenter(die, pad) {
  const rotated = transformPoint(pad.x, pad.y, die.diesize.x, die.diesize.y, die.rotation);
  return { x: rotated.x + die.posX, y: rotated.y + die.posY };
}

function getPadWorldRect(die, pad) {
  const center = getPadWorldCenter(die, pad);
  const rotatedSize = getRotatedRectSize(pad.width, pad.height, die.rotation);
  return {
    left: center.x - rotatedSize.width / 2,
    right: center.x + rotatedSize.width / 2,
    top: center.y + rotatedSize.height / 2,
    bottom: center.y - rotatedSize.height / 2,
    cx: center.x,
    cy: center.y,
  };
}

function unionRects(rects) {
  return rects.reduce(
    (acc, r) => ({
      minX: Math.min(acc.minX, r.minX),
      maxX: Math.max(acc.maxX, r.maxX),
      minY: Math.min(acc.minY, r.minY),
      maxY: Math.max(acc.maxY, r.maxY),
    }),
    { minX: Infinity, maxX: -Infinity, minY: Infinity, maxY: -Infinity },
  );
}

function expandRect(rect, margin) {
  return { minX: rect.minX - margin, maxX: rect.maxX + margin, minY: rect.minY - margin, maxY: rect.maxY + margin };
}

function defaultDiePos(dieIndex) {
  // 預設與 dies[0] 中心對齊
  const anchor = state.dies[0];
  const die = state.dies[dieIndex];
  if (!anchor || !die) return { x: 0, y: 0 };
  const anchorCx = anchor.posX + anchor.diesize.x / 2;
  const anchorCy = anchor.posY + anchor.diesize.y / 2;
  return { x: anchorCx - die.diesize.x / 2, y: anchorCy - die.diesize.y / 2 };
}

function computeDefaultPositions() {
  if (!state.dies.length) return;
  state.dies[0].posX = 0;
  state.dies[0].posY = 0;
  for (let i = 1; i < state.dies.length; i++) {
    const def = defaultDiePos(i);
    state.dies[i].posX = def.x;
    state.dies[i].posY = def.y;
  }
}

function getDieClusterBounds() {
  if (!state.dies.length) return null;
  return unionRects(state.dies.map((die) => getDieWorldRect(die)));
}

function getRails() {
  const cluster = getDieClusterBounds();
  if (!cluster) return null;
  const inner = expandRect(cluster, RAIL_INNER_MARGIN);
  if (state.dies.length <= 1) {
    // 單晶片：只放一個方形線條
    return { inner, outer: inner, single: true };
  }
  const outer = expandRect(inner, RAIL_GAP);
  return { inner, outer, single: false };
}

function ensureSipDefaults() {
  const rails = getRails();
  if (!rails) return;
  if (state.sip.auto) {
    const auto = expandRect(rails.outer, SIP_AUTO_MARGIN);
    state.sip.posX = auto.minX;
    state.sip.posY = auto.minY;
    state.sip.sizeX = auto.maxX - auto.minX;
    state.sip.sizeY = auto.maxY - auto.minY;
  }
}

function getSipRect() {
  return {
    minX: state.sip.posX,
    minY: state.sip.posY,
    maxX: state.sip.posX + state.sip.sizeX,
    maxY: state.sip.posY + state.sip.sizeY,
  };
}

// ── Wirebond 解析與路由 ──────────────────────────────────────

function classifyWirebondToken(token) {
  const t = String(token).trim();
  if (!t) return { type: 'none' };
  if (/^nc$/i.test(t)) return { type: 'none' };
  if (/^lead\s*frame$/i.test(t)) return { type: 'gnd' };
  if (/^D\d+\.[\w-]+$/i.test(t)) return { type: 'dieref', ref: t };
  return { type: 'io', io: t };
}

function buildGlobalPadIndex() {
  const map = new Map();
  state.dies.forEach((die, dieIndex) => {
    die.pads.forEach((pad) => {
      map.set(String(pad.label), { dieIndex, pad });
    });
  });
  return map;
}

function collectWirebondTasks() {
  const padIndex = buildGlobalPadIndex();
  const bonds = []; // { aDieIndex, aPad, bDieIndex, bPad, sameDie }
  const breakouts = []; // { dieIndex, pad, type: 'gnd'|'io', ioLabel }
  const bondKeys = new Set();

  state.dies.forEach((die, dieIndex) => {
    die.pads.forEach((pad) => {
      if (pad.wirebond == null || String(pad.wirebond).trim() === '') return;
      const tokens = [...new Set(String(pad.wirebond).split(',').map((s) => s.trim()).filter(Boolean))];
      let hasGnd = false;
      const ioLabels = new Set();

      tokens.forEach((token) => {
        const cls = classifyWirebondToken(token);
        if (cls.type === 'none') return;
        if (cls.type === 'gnd') {
          hasGnd = true;
          return;
        }
        if (cls.type === 'io') {
          ioLabels.add(cls.io);
          return;
        }
        if (cls.type === 'dieref') {
          const target = padIndex.get(cls.ref);
          if (!target) {
            // 找不到目標（例如單晶片模式下參照其他 Die）：往外接腳位並標註目標
            breakouts.push({ dieIndex, pad, type: 'dieref', ioLabel: cls.ref });
            return;
          }
          const key = [`${dieIndex}:${pad.label}`, `${target.dieIndex}:${target.pad.label}`].sort().join('|');
          if (bondKeys.has(key)) return;
          bondKeys.add(key);
          bonds.push({
            aDieIndex: dieIndex,
            aPad: pad,
            bDieIndex: target.dieIndex,
            bPad: target.pad,
            sameDie: dieIndex === target.dieIndex,
          });
        }
      });

      if (hasGnd) breakouts.push({ dieIndex, pad, type: 'gnd', ioLabel: 'GND' });
      ioLabels.forEach((io) => breakouts.push({ dieIndex, pad, type: 'io', ioLabel: io }));
    });
  });

  return { bonds, breakouts };
}

function nearestEdge(point, rect) {
  const distances = {
    top: Math.abs(rect.maxY - point.y),
    bottom: Math.abs(point.y - rect.minY),
    left: Math.abs(point.x - rect.minX),
    right: Math.abs(rect.maxX - point.x),
  };
  return Object.entries(distances).sort((a, b) => a[1] - b[1])[0][0];
}

// 依腳位在「自己晶片」的哪一側決定出線方向：
// 上側腳位往上、下側往下、左側往左、右側往右，同一側的腳位往同方向規劃。
function padSideOfDie(die, pad) {
  const dieRect = getDieWorldRect(die);
  const center = getPadWorldCenter(die, pad);
  return nearestEdge(center, dieRect);
}

function slotPointOnEdge(rect, edge, t, inset) {
  const clampT = Math.min(Math.max(t, 0), 1);
  switch (edge) {
    case 'top':
      return { x: rect.minX + inset + clampT * (rect.maxX - rect.minX - 2 * inset), y: rect.maxY };
    case 'bottom':
      return { x: rect.minX + inset + clampT * (rect.maxX - rect.minX - 2 * inset), y: rect.minY };
    case 'left':
      return { x: rect.minX, y: rect.minY + inset + clampT * (rect.maxY - rect.minY - 2 * inset) };
    case 'right':
    default:
      return { x: rect.maxX, y: rect.minY + inset + clampT * (rect.maxY - rect.minY - 2 * inset) };
  }
}

function assignBreakoutSlots(breakouts, rails) {
  // GND -> 內圈線 (inner rail) ； IO / 跨Die參照 -> 外圈線 (outer rail)，分層降低交叉
  // 單晶片模式（rails.single）：全部放在同一個方形線上，同邊合併排序避免重疊
  const groups = new Map(); // key -> [{breakout, padCenter, along}]

  breakouts.forEach((b) => {
    const rail = b.type === 'gnd' ? rails.inner : rails.outer;
    const die = state.dies[b.dieIndex];
    const padCenter = getPadWorldCenter(die, b.pad);
    // 出線方向由腳位在晶片的哪一側決定（上/下/左/右同側同方向）
    const edge = padSideOfDie(die, b.pad);
    const along = edge === 'top' || edge === 'bottom' ? padCenter.x : padCenter.y;
    const key = rails.single ? edge : `${b.type === 'gnd' ? 'inner' : 'outer'}|${edge}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ breakout: b, padCenter, along, rail, edge });
  });

  const results = [];
  groups.forEach((items) => {
    items.sort((a, b) => a.along - b.along);
    const n = items.length;
    items.forEach((item, i) => {
      const t = n === 1 ? 0.5 : i / (n - 1);
      const slot = slotPointOnEdge(item.rail, item.edge, t, 400);
      results.push({ ...item.breakout, padCenter: item.padCenter, slot, edge: item.edge });
    });
  });
  return results;
}

// ── 畫面繪製 ────────────────────────────────────────────────

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

function getSceneBounds() {
  if (!state.dies.length) return null;
  const dieRects = state.dies.map((die) => getDieWorldRect(die));
  const rails = getRails();
  const sipRect = getSipRect();
  const rects = [...dieRects, rails.outer, sipRect];
  return unionRects(rects);
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

function drawAxes(bounds, mapX, mapY, dpr) {
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

function drawRect(rect, mapX, mapY, strokeStyle, dashed) {
  context.save();
  context.strokeStyle = strokeStyle;
  context.lineWidth = 1.5;
  if (dashed) context.setLineDash([10, 6]);
  context.strokeRect(mapX(rect.minX), mapY(rect.maxY), mapX(rect.maxX) - mapX(rect.minX), mapY(rect.minY) - mapY(rect.maxY));
  context.restore();
}

function drawSipOutline(mapX, mapY, dpr) {
  const rect = getSipRect();
  context.save();
  context.strokeStyle = '#334155';
  context.lineWidth = 2.5;
  context.setLineDash([]);
  context.strokeRect(mapX(rect.minX), mapY(rect.maxY), mapX(rect.maxX) - mapX(rect.minX), mapY(rect.minY) - mapY(rect.maxY));
  context.fillStyle = '#334155';
  context.font = `bold ${14 * dpr * state.fontScale}px Segoe UI`;
  context.fillText('SiP Outline', mapX(rect.minX), mapY(rect.maxY) - 10 * dpr);
  context.restore();
}

function drawRails(rails, mapX, mapY) {
  if (rails.single) {
    drawRect(rails.inner, mapX, mapY, '#2563eb99', true);
    return;
  }
  drawRect(rails.inner, mapX, mapY, '#16a34a99', true);
  drawRect(rails.outer, mapX, mapY, '#2563eb99', true);
}

function drawDie(die, dieIndex, mapX, mapY, scale, dpr) {
  const worldRect = getDieWorldRect(die);
  const isActive = dieIndex === state.activeDieIndex && state.dies.length > 1;
  context.save();
  context.beginPath();
  context.rect(mapX(worldRect.minX), mapY(worldRect.maxY), mapX(worldRect.maxX) - mapX(worldRect.minX), mapY(worldRect.minY) - mapY(worldRect.maxY));
  context.fillStyle = isActive ? 'rgba(249, 115, 22, 0.08)' : 'rgba(37, 99, 235, 0.06)';
  context.strokeStyle = isActive ? '#f97316' : '#2563eb';
  context.lineWidth = isActive ? 3 : 2;
  context.fill();
  context.stroke();

  const fontSize = Math.round(16 * dpr * state.fontScale);
  context.fillStyle = isActive ? '#c2410c' : '#1d4ed8';
  context.font = `bold ${fontSize}px Segoe UI`;
  context.textAlign = 'start';
  const label = `${die.diename}${die.rotation ? ` (${die.rotation}°)` : ''}`;
  context.fillText(label, mapX(worldRect.minX), mapY(worldRect.maxY) - 8 * dpr);
  context.restore();

  // Pads
  const drawLabels = scale >= 0.12;
  const padFontSize = Math.round(11 * dpr * state.fontScale);
  context.save();
  context.font = `${padFontSize}px Consolas`;
  die.pads.forEach((pad) => {
    const r = getPadWorldRect(die, pad);
    const left = mapX(r.left);
    const right = mapX(r.right);
    const top = mapY(r.top);
    const bottom = mapY(r.bottom);
    context.fillStyle = '#0ea5e9';
    context.strokeStyle = '#0369a1';
    context.lineWidth = 1;
    context.fillRect(left, top, right - left, bottom - top);
    context.strokeRect(left, top, right - left, bottom - top);
    if (drawLabels) {
      context.fillStyle = '#0f172a';
      context.fillText(String(pad.label), left + 3, top + 12 * dpr);
    }
  });
  context.restore();
}

function drawBonds(bonds, mapX, mapY, dpr) {
  context.save();
  context.lineWidth = 1.8;
  bonds.forEach((bond) => {
    const a = getPadWorldCenter(state.dies[bond.aDieIndex], bond.aPad);
    const b = getPadWorldCenter(state.dies[bond.bDieIndex], bond.bPad);
    context.strokeStyle = bond.sameDie ? '#0ea5e9' : '#f97316';
    context.beginPath();
    context.moveTo(mapX(a.x), mapY(a.y));
    context.lineTo(mapX(b.x), mapY(b.y));
    context.stroke();
  });
  context.restore();
}

function drawBreakouts(resolvedBreakouts, mapX, mapY, dpr) {
  // 判斷哪些 IO label 屬於共接（多個來源腳位）
  const ioGroups = new Map();
  resolvedBreakouts.forEach((b) => {
    if (b.type !== 'io') return;
    if (!ioGroups.has(b.ioLabel)) ioGroups.set(b.ioLabel, []);
    ioGroups.get(b.ioLabel).push(b);
  });

  context.save();
  // 連線（pad -> slot）
  resolvedBreakouts.forEach((b) => {
    context.strokeStyle = b.type === 'gnd' ? '#16a34a' : b.type === 'dieref' ? '#f97316' : '#2563eb';
    context.lineWidth = 1.4;
    context.beginPath();
    context.moveTo(mapX(b.padCenter.x), mapY(b.padCenter.y));
    context.lineTo(mapX(b.slot.x), mapY(b.slot.y));
    context.stroke();
  });

  // 共接 IO 之間的紅色連接線
  ioGroups.forEach((items) => {
    if (items.length < 2) return;
    context.strokeStyle = '#dc2626';
    context.lineWidth = 1.2;
    context.setLineDash([4, 4]);
    context.beginPath();
    items.forEach((item, i) => {
      const p = item.slot;
      if (i === 0) context.moveTo(mapX(p.x), mapY(p.y));
      else context.lineTo(mapX(p.x), mapY(p.y));
    });
    context.stroke();
    context.setLineDash([]);
  });

  // 標註方框與文字
  const boxFontSize = Math.round(11 * dpr * state.fontScale);
  context.font = `bold ${boxFontSize}px Segoe UI`;
  resolvedBreakouts.forEach((b) => {
    const isSharedIo = b.type === 'io' && ioGroups.get(b.ioLabel).length > 1;
    const boxColor = b.type === 'gnd' ? '#16a34a' : b.type === 'dieref' ? '#f97316' : '#2563eb';
    const sx = mapX(b.slot.x);
    const sy = mapY(b.slot.y);
    const boxW = 10 * dpr;
    const boxH = 7 * dpr;
    context.fillStyle = boxColor;
    context.fillRect(sx - boxW / 2, sy - boxH / 2, boxW, boxH);
    context.fillStyle = isSharedIo ? '#dc2626' : b.type === 'dieref' ? '#c2410c' : '#0f172a';
    // IO 外接腳位標註：使用原本 Die 的 name 欄位；dieref 標註目標腳位
    const text = b.type === 'gnd' ? 'GND'
      : b.type === 'dieref' ? `→${b.ioLabel}`
      : (String(b.pad.name || '').trim() || b.ioLabel);
    let dx = 8 * dpr;
    let dy = 4 * dpr;
    if (b.edge === 'top') dy = -8 * dpr;
    if (b.edge === 'left') dx = -8 * dpr - context.measureText(text).width;
    context.fillText(text, sx + dx, sy + dy);
  });

  // 共接 IO：紅字標註這幾個點連在一起（列出各腳位 name）
  context.font = `bold ${Math.round(12 * dpr * state.fontScale)}px Segoe UI`;
  context.fillStyle = '#dc2626';
  ioGroups.forEach((items, ioLabel) => {
    if (items.length < 2) return;
    const names = items.map((it) => String(it.pad.name || it.pad.label || '').trim()).join(' + ');
    // 標註放在群組最後一個 slot 外側
    const last = items[items.length - 1];
    const sx = mapX(last.slot.x);
    const sy = mapY(last.slot.y);
    let dx = 8 * dpr;
    let dy = 18 * dpr;
    if (last.edge === 'top') dy = -20 * dpr;
    if (last.edge === 'left') dx = -8 * dpr - context.measureText(`${names} 共接`).width;
    context.fillText(`${names} 共接 (IO ${ioLabel})`, sx + dx, sy + dy);
  });
  context.restore();
}

// ── 腳位選取與 Highlight ──────────────────────────────────────

const HIGHLIGHT_COLOR = '#ff00e5'; // 螢光桃紅
const HIGHLIGHT_GLOW = 'rgba(255, 0, 229, 0.65)';

function samePad(a, dieIndex, pad) {
  return a && a.dieIndex === dieIndex && String(a.label) === String(pad.label);
}

function findPadAt(worldPoint) {
  for (let i = state.dies.length - 1; i >= 0; i--) {
    const die = state.dies[i];
    for (const pad of die.pads) {
      const r = getPadWorldRect(die, pad);
      const minX = Math.min(r.left, r.right);
      const maxX = Math.max(r.left, r.right);
      const minY = Math.min(r.bottom, r.top);
      const maxY = Math.max(r.bottom, r.top);
      // 留一點點擊容差（世界座標 10um）
      const tol = 10;
      if (worldPoint.x >= minX - tol && worldPoint.x <= maxX + tol && worldPoint.y >= minY - tol && worldPoint.y <= maxY + tol) {
        return { dieIndex: i, pad };
      }
    }
  }
  return null;
}

// 回傳選取腳位的所有連線與連接腳位
 function getPadConnections(sel, bonds, resolvedBreakouts) {
  const relatedBonds = [];
  const connectedPads = []; // { dieIndex, pad }
  bonds.forEach((bond) => {
    const isA = bond.aDieIndex === sel.dieIndex && String(bond.aPad.label) === String(sel.label);
    const isB = bond.bDieIndex === sel.dieIndex && String(bond.bPad.label) === String(sel.label);
    if (isA || isB) {
      relatedBonds.push(bond);
      connectedPads.push(isA
        ? { dieIndex: bond.bDieIndex, pad: bond.bPad }
        : { dieIndex: bond.aDieIndex, pad: bond.aPad });
    }
  });
  const relatedBreakouts = resolvedBreakouts.filter(
    (b) => b.dieIndex === sel.dieIndex && String(b.pad.label) === String(sel.label),
  );
  return { relatedBonds, connectedPads, relatedBreakouts };
}

function drawSelectionHighlight(bonds, resolvedBreakouts, mapX, mapY, dpr) {
  const sel = state.selectedPad;
  if (!sel) return;
  const die = state.dies[sel.dieIndex];
  if (!die) return;
  const pad = die.pads.find((p) => String(p.label) === String(sel.label));
  if (!pad) return;

  const { relatedBonds, connectedPads, relatedBreakouts } = getPadConnections(sel, bonds, resolvedBreakouts);

  context.save();
  context.shadowColor = HIGHLIGHT_GLOW;
  context.shadowBlur = 12 * dpr;

  // Highlight 連線（bond）
  context.strokeStyle = HIGHLIGHT_COLOR;
  context.lineWidth = 3.5;
  relatedBonds.forEach((bond) => {
    const a = getPadWorldCenter(state.dies[bond.aDieIndex], bond.aPad);
    const b = getPadWorldCenter(state.dies[bond.bDieIndex], bond.bPad);
    context.beginPath();
    context.moveTo(mapX(a.x), mapY(a.y));
    context.lineTo(mapX(b.x), mapY(b.y));
    context.stroke();
  });

  // Highlight 外接線（breakout）
  relatedBreakouts.forEach((b) => {
    context.beginPath();
    context.moveTo(mapX(b.padCenter.x), mapY(b.padCenter.y));
    context.lineTo(mapX(b.slot.x), mapY(b.slot.y));
    context.stroke();
    // slot 端點框
    const sx = mapX(b.slot.x);
    const sy = mapY(b.slot.y);
    context.strokeRect(sx - 7 * dpr, sy - 5 * dpr, 14 * dpr, 10 * dpr);
  });

  // Highlight 被選腳位與所有連接腳位
  const highlightPadRect = (dIdx, p) => {
    const d = state.dies[dIdx];
    const r = getPadWorldRect(d, p);
    const left = mapX(r.left);
    const right = mapX(r.right);
    const top = mapY(r.top);
    const bottom = mapY(r.bottom);
    context.lineWidth = 3;
    context.strokeStyle = HIGHLIGHT_COLOR;
    context.strokeRect(left - 2, top - 2, (right - left) + 4, (bottom - top) + 4);
  };
  highlightPadRect(sel.dieIndex, pad);
  connectedPads.forEach((c) => highlightPadRect(c.dieIndex, c.pad));

  context.restore();

  // 資訊方塊
  const dieTag = (i) => `D${i + 1}`;
  const lines = [`${dieTag(sel.dieIndex)}.${pad.label}  ${pad.name || ''}`.trim()];
  if (connectedPads.length) {
    lines.push('─ 連接腳位 ─');
    connectedPads.forEach((c) => {
      lines.push(`${dieTag(c.dieIndex)}.${c.pad.label}  ${c.pad.name || ''}`.trim());
    });
  }
  relatedBreakouts.forEach((b) => {
    lines.push(
      b.type === 'gnd' ? '─ 外接：GND (lead frame)'
      : b.type === 'dieref' ? `─ 外接至其他 Die：${b.ioLabel}`
      : `─ 外接 IO：${String(b.pad.name || '').trim() || b.ioLabel}`,
    );
  });
  if (connectedPads.length === 0 && relatedBreakouts.length === 0) {
    const wb = String(pad.wirebond ?? '').trim();
    lines.push(wb.toUpperCase() === 'NC' ? '─ NC（空接）' : '─ 無連線資訊');
  }

  // 畫在腳位旁邊
  const c = getPadWorldCenter(die, pad);
  const anchorX = mapX(c.x) + 14 * dpr;
  const anchorY = mapY(c.y) - 14 * dpr;
  const fontSize = Math.round(12 * dpr * state.tooltipFontScale);
  context.save();
  context.font = `${fontSize}px Consolas, monospace`;
  const lineH = fontSize * 1.35;
  const boxW = Math.max(...lines.map((t) => context.measureText(t).width)) + 16 * dpr;
  const boxH = lines.length * lineH + 12 * dpr;
  let bx = anchorX;
  let by = anchorY - boxH;
  // 避免超出畫布上緣/右緣
  if (by < 4) by = anchorY + 8 * dpr;
  if (bx + boxW > canvas.width - 4) bx = mapX(c.x) - boxW - 14 * dpr;
  context.fillStyle = 'rgba(15, 23, 42, 0.92)';
  context.strokeStyle = HIGHLIGHT_COLOR;
  context.lineWidth = 2;
  context.beginPath();
  context.roundRect(bx, by, boxW, boxH, 8 * dpr);
  context.fill();
  context.stroke();
  context.fillStyle = '#ffffff';
  lines.forEach((t, i) => {
    if (t.startsWith('─')) context.fillStyle = '#7dd3fc';
    else if (i === 0) context.fillStyle = '#f0abfc';
    else context.fillStyle = '#ffffff';
    context.fillText(t, bx + 8 * dpr, by + 8 * dpr + (i + 0.8) * lineH - lineH * 0.35);
  });
  context.restore();
}

function updateSummary() {
  const totalPads = state.dies.reduce((sum, d) => sum + d.pads.length, 0);
  padCountValue.textContent = state.dies.length ? String(totalPads) : '-';
  const activeDie = state.dies[state.activeDieIndex];
  rotationValue.textContent = activeDie ? `${activeDie.rotation}°` : '0°';
  zoomValue.textContent = `${Math.round(state.zoomMultiplier * 100)}%`;
  sourceValue.textContent = state.sourceName;
}

function updateDieListSummary() {
  if (!dieListSummary) return;
  dieListSummary.innerHTML = '';
  state.dies.forEach((die, i) => {
    const div = document.createElement('div');
    const dt = document.createElement('dt');
    dt.textContent = `${i === state.activeDieIndex ? '▶ ' : ''}${die.diename}`;
    const dd = document.createElement('dd');
    dd.textContent = `${die.pads.length} pads / ${die.rotation}°`;
    div.appendChild(dt);
    div.appendChild(dd);
    dieListSummary.appendChild(div);
  });
}

function updateActiveDieSelect() {
  if (!activeDieSelect) return;
  activeDieSelect.innerHTML = '';
  state.dies.forEach((die, i) => {
    const opt = document.createElement('option');
    opt.value = String(i);
    opt.textContent = `${i === 0 ? 'D1 ' : i === 1 ? 'D2 ' : `D${i + 1} `}${die.diename}`;
    activeDieSelect.appendChild(opt);
  });
  activeDieSelect.value = String(state.activeDieIndex);
}

function updateSipInputs() {
  if (state.mode !== 'sip') return;
  if (state.dies.length > 1) {
    const def = defaultDiePos(1);
    const offX = round3(state.dies[1].posX - def.x);
    const offY = round3(state.dies[1].posY - def.y);
    die2OffsetXInput.value = offX;
    die2OffsetYInput.value = offY;
  }
  sipSizeXInput.value = round3(state.sip.sizeX);
  sipSizeYInput.value = round3(state.sip.sizeY);
}

function centerView() {
  if (!state.dies.length) { state.panX = 0; state.panY = 0; return; }
  const rect = canvas.getBoundingClientRect();
  const cssW = rect.width || canvas.width / (window.devicePixelRatio || 1);
  const cssH = rect.height || canvas.height / (window.devicePixelRatio || 1);
  const bounds = getSceneBounds();
  const sceneW = Math.max(bounds.maxX - bounds.minX, 1);
  const sceneH = Math.max(bounds.maxY - bounds.minY, 1);
  const PAD = 60;
  const fitScale = Math.min((cssW - PAD * 2) / sceneW, (cssH - PAD * 2) / sceneH);
  const sc = fitScale * state.zoomMultiplier;
  state.panX = PAD + ((cssW - PAD * 2) - sceneW * sc) / 2 - bounds.minX * sc;
  state.panY = PAD + ((cssH - PAD * 2) - sceneH * sc) / 2 + bounds.maxY * sc;
}

let lastMapX = (x) => x;
let lastMapY = (y) => y;
let lastScale = 1;

function drawScene() {
  resizeCanvasForDisplay();
  const dpr = window.devicePixelRatio || 1;
  const width = canvas.width;
  const height = canvas.height;
  const cssW = width / dpr;
  const cssH = height / dpr;

  context.clearRect(0, 0, width, height);
  context.fillStyle = '#f8fbff';
  context.fillRect(0, 0, width, height);

  if (!state.dies.length) {
    context.fillStyle = '#64748b';
    context.font = `${18 * dpr}px Segoe UI`;
    context.fillText('請先匯入 Excel 或 JSON。', 48, 64);
    updateSummary();
    return;
  }

  ensureSipDefaults();

  const bounds = getSceneBounds();
  const sceneW = Math.max(bounds.maxX - bounds.minX, 1);
  const sceneH = Math.max(bounds.maxY - bounds.minY, 1);
  const PAD = 60;
  const fitScale = Math.min((cssW - PAD * 2) / sceneW, (cssH - PAD * 2) / sceneH);
  const sc = fitScale * state.zoomMultiplier;

  const mapX = (x) => (state.panX + x * sc) * dpr;
  const mapY = (y) => (state.panY - y * sc) * dpr;
  lastMapX = mapX; lastMapY = mapY; lastScale = sc;

  drawGrid(bounds, mapX, mapY, sc * dpr);
  drawAxes(bounds, mapX, mapY, dpr);

  // 只要有 wirebond 資料（單晶片或 SiP）都繪製打線
  const { bonds, breakouts } = collectWirebondTasks();
  const hasWirebond = bonds.length > 0 || breakouts.length > 0;

  if (state.mode === 'sip') {
    drawSipOutline(mapX, mapY, dpr);
  }
  if (state.mode === 'sip' || hasWirebond) {
    const rails = getRails();
    drawRails(rails, mapX, mapY);
  }

  state.dies.forEach((die, i) => drawDie(die, i, mapX, mapY, sc * dpr, dpr));

  if (hasWirebond) {
    const rails = getRails();
    const resolvedBreakouts = assignBreakoutSlots(breakouts, rails);
    drawBonds(bonds, mapX, mapY, dpr);
    drawBreakouts(resolvedBreakouts, mapX, mapY, dpr);
    drawSelectionHighlight(bonds, resolvedBreakouts, mapX, mapY, dpr);
  }

  updateSummary();
  updateDieListSummary();
}

// ── 狀態載入／套用 ──────────────────────────────────────────

function syncInputsFromState() {
  const die = state.dies[state.activeDieIndex];
  if (!die) {
    dieNameInput.value = '';
    dieSizeXInput.value = '';
    dieSizeYInput.value = '';
    return;
  }
  dieNameInput.value = die.diename;
  dieSizeXInput.value = die.diesize.x;
  dieSizeYInput.value = die.diesize.y;
}

function syncStateFromInputs() {
  const die = state.dies[state.activeDieIndex];
  if (!die) return;
  const name = dieNameInput.value.trim();
  const sizeX = Number(dieSizeXInput.value);
  const sizeY = Number(dieSizeYInput.value);
  if (name) die.diename = name;
  if (Number.isFinite(sizeX) && sizeX > 0) die.diesize.x = sizeX;
  if (Number.isFinite(sizeY) && sizeY > 0) die.diesize.y = sizeY;
}

function applyParsedResult(payload, sourceName, warningOverride) {
  if (payload.sipAsset) {
    state.mode = 'sip';
    state.dies = payload.sipAsset.dies.map((d) => ({
      sheet: d.sheet,
      diename: d.diename,
      diesize: { x: Number(d.diesize.x), y: Number(d.diesize.y) },
      pads: d.pads.map((p) => ({ ...p })),
      rotation: 0,
      posX: 0,
      posY: 0,
    }));
    computeDefaultPositions();
    state.sip = { sizeX: 0, sizeY: 0, posX: 0, posY: 0, auto: true };
    state.activeDieIndex = Math.min(1, state.dies.length - 1);
    sipPanel.style.display = '';
  } else if (payload.dieAsset || payload.die) {
    const asset = payload.dieAsset || payload;
    state.mode = 'single';
    state.dies = [{
      sheet: asset.die ? asset.die.diename : asset.diename,
      diename: asset.die ? asset.die.diename : asset.diename,
      diesize: asset.die ? { ...asset.die.diesize } : { ...asset.diesize },
      pads: (asset.pads || []).map((p) => ({ ...p })),
      rotation: 0,
      posX: 0,
      posY: 0,
    }];
    state.activeDieIndex = 0;
    sipPanel.style.display = 'none';
  } else {
    throw new Error('無法辨識的資料格式。');
  }

  state.sourceName = sourceName;
  state.zoomMultiplier = 1;
  syncInputsFromState();
  updateActiveDieSelect();
  centerView();
  updateSipInputs();
  drawScene();
  const warning = warningOverride ?? payload.warning ?? '';
  setStatus(warning ? `載入完成：${sourceName}\n${warning}` : `載入完成：${sourceName}`, warning ? 'warning' : 'success');
}

async function handleExcelFile(file) {
  try {
    const buffer = await readFileAsArrayBuffer(file);
    const result = await parseExcelWithFallback(buffer, file.name);
    applyParsedResult(result, file.name);
    if (result.parser) {
      setStatus(`${file.name} 載入完成\n解析模式：${result.parser}${result.warning ? `\n${result.warning}` : ''}`, result.warning ? 'warning' : 'success');
    }
  } catch (error) {
    setStatus(error.message, 'error');
  }
}

async function handleJsonFile(file) {
  try {
    const text = await readFileAsText(file);
    const parsed = JSON.parse(text);
    if (parsed.dies && Array.isArray(parsed.dies)) {
      state.mode = 'sip';
      state.dies = parsed.dies.map((d) => ({
        sheet: d.sheet,
        diename: d.diename,
        diesize: { ...d.diesize },
        pads: (d.pads || []).map((p) => ({ ...p })),
        rotation: d.rotation || 0,
        posX: d.pos ? d.pos.x : 0,
        posY: d.pos ? d.pos.y : 0,
      }));
      state.activeDieIndex = Math.min(1, state.dies.length - 1);
      state.sip = parsed.sip
        ? { sizeX: parsed.sip.size.x, sizeY: parsed.sip.size.y, posX: parsed.sip.pos.x, posY: parsed.sip.pos.y, auto: false }
        : { sizeX: 0, sizeY: 0, posX: 0, posY: 0, auto: true };
      sipPanel.style.display = '';
      state.sourceName = file.name;
      state.zoomMultiplier = 1;
      syncInputsFromState();
      updateActiveDieSelect();
      centerView();
      updateSipInputs();
      drawScene();
      setStatus(`載入完成：${file.name}`, 'success');
    } else {
      applyParsedResult({ dieAsset: parsed }, file.name, '');
    }
  } catch (error) {
    setStatus(`JSON 載入失敗：${error.message}`, 'error');
  }
}

function exportDieAssetJson() {
  try {
    if (!state.dies.length) throw new Error('目前沒有可匯出的 Die 資料。');
    syncStateFromInputs();

    let blob;
    let defaultName;
    if (state.mode === 'sip') {
      const payload = {
        format_version: '4.0',
        unit: 'um',
        sip: { size: { x: round3(state.sip.sizeX), y: round3(state.sip.sizeY) }, pos: { x: round3(state.sip.posX), y: round3(state.sip.posY) } },
        dies: state.dies.map((d) => ({
          sheet: d.sheet,
          diename: d.diename,
          diesize: { x: round3(d.diesize.x), y: round3(d.diesize.y) },
          pads: d.pads,
          rotation: d.rotation,
          pos: { x: round3(d.posX), y: round3(d.posY) },
        })),
      };
      blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      defaultName = 'sip_layout';
    } else {
      const die = state.dies[0];
      const payload = {
        format_version: '3.0',
        unit: 'um',
        die: { diename: die.diename, diesize: { x: round3(die.diesize.x), y: round3(die.diesize.y) } },
        pads: die.pads,
      };
      blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
      defaultName = die.diename || 'die-asset';
    }

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    const safeName = String(defaultName).replace(/[^a-z0-9\-_]+/gi, '_');
    anchor.href = url;
    anchor.download = `${safeName}_asset.json`;
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
    if (!response.ok) throw new Error('無法透過本機伺服器讀取範例檔。');
    const payload = await response.json();
    if (!payload?.dieAsset && !payload?.sipAsset) throw new Error(payload.error || '範例解析失敗。');
    applyParsedResult(payload, 'GBB898die.xlsx', payload.warning || '');
  } catch (error) {
    setStatus(`載入範例失敗：${error.message}`, 'error');
  }
}

// ── 事件綁定 ────────────────────────────────────────────────

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
    const die = state.dies[state.activeDieIndex];
    if (!die) return;
    die.rotation = rotation;
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

if (activeDieSelect) {
  activeDieSelect.addEventListener('change', () => {
    state.activeDieIndex = Number(activeDieSelect.value);
    syncInputsFromState();
    drawScene();
  });
}

if (die2OffsetXInput) {
  die2OffsetXInput.addEventListener('change', () => {
    if (state.dies.length < 2) return;
    const def = defaultDiePos(1);
    const offX = Number(die2OffsetXInput.value) || 0;
    state.dies[1].posX = def.x + offX;
    drawScene();
  });
}
if (die2OffsetYInput) {
  die2OffsetYInput.addEventListener('change', () => {
    if (state.dies.length < 2) return;
    const def = defaultDiePos(1);
    const offY = Number(die2OffsetYInput.value) || 0;
    state.dies[1].posY = def.y + offY;
    drawScene();
  });
}
if (sipSizeXInput) {
  sipSizeXInput.addEventListener('change', () => {
    const val = Number(sipSizeXInput.value);
    if (!Number.isFinite(val) || val <= 0) return;
    state.sip.sizeX = val;
    state.sip.auto = false;
    drawScene();
  });
}
if (sipSizeYInput) {
  sipSizeYInput.addEventListener('change', () => {
    const val = Number(sipSizeYInput.value);
    if (!Number.isFinite(val) || val <= 0) return;
    state.sip.sizeY = val;
    state.sip.auto = false;
    drawScene();
  });
}

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

// 資訊方塊字體大小按鈕
const tooltipFontUpButton = document.getElementById('tooltipFontUpButton');
const tooltipFontDownButton = document.getElementById('tooltipFontDownButton');
const tooltipFontScaleValue = document.getElementById('tooltipFontScaleValue');

function updateTooltipFontScaleLabel() {
  if (tooltipFontScaleValue) tooltipFontScaleValue.textContent = `${Math.round(state.tooltipFontScale * 100)}%`;
}

if (tooltipFontUpButton) {
  tooltipFontUpButton.addEventListener('click', () => {
    state.tooltipFontScale = Math.min(state.tooltipFontScale * 1.25, 8);
    updateTooltipFontScaleLabel();
    drawScene();
  });
}
if (tooltipFontDownButton) {
  tooltipFontDownButton.addEventListener('click', () => {
    state.tooltipFontScale = Math.max(state.tooltipFontScale / 1.25, 0.3);
    updateTooltipFontScaleLabel();
    drawScene();
  });
}
updateTooltipFontScaleLabel();

// ── 拖拉：晶片 / SiP 外框 / 畫布平移 ─────────────────────────

function screenToWorld(clientX, clientY) {
  const rect = canvas.getBoundingClientRect();
  const cssX = clientX - rect.left;
  const cssY = clientY - rect.top;
  const dpr = window.devicePixelRatio || 1;
  // mapX(x) = (panX + x*sc) * dpr  → solve x
  const x = ((cssX * dpr) / dpr - state.panX) / lastScale;
  const y = (state.panY - (cssY * dpr) / dpr) / lastScale;
  return { x, y };
}

function rectContains(rect, point) {
  return point.x >= rect.minX && point.x <= rect.maxX && point.y >= rect.minY && point.y <= rect.maxY;
}

let drag = null;
let pendingClick = null; // 偵測 click（滑鼠沒有移動）用

canvas.addEventListener('mousedown', (e) => {
  if (e.button !== 0) return;
  const worldPoint = screenToWorld(e.clientX, e.clientY);
  pendingClick = { clientX: e.clientX, clientY: e.clientY, worldPoint };

  let hitDieIndex = -1;
  for (let i = state.dies.length - 1; i >= 0; i--) {
    const rect = getDieWorldRect(state.dies[i]);
    if (rectContains(rect, worldPoint)) { hitDieIndex = i; break; }
  }

  if (hitDieIndex >= 0) {
    const die = state.dies[hitDieIndex];
    state.activeDieIndex = hitDieIndex;
    updateActiveDieSelect();
    syncInputsFromState();
    drag = { type: 'die', index: hitDieIndex, startClientX: e.clientX, startClientY: e.clientY, startPosX: die.posX, startPosY: die.posY };
  } else if (state.mode === 'sip' && rectContains(getSipRect(), worldPoint)) {
    drag = { type: 'sip', startClientX: e.clientX, startClientY: e.clientY, startPosX: state.sip.posX, startPosY: state.sip.posY };
  } else {
    drag = { type: 'pan', startClientX: e.clientX, startClientY: e.clientY, panX: state.panX, panY: state.panY };
  }
  canvas.style.cursor = 'grabbing';
  e.preventDefault();
});

window.addEventListener('mousemove', (e) => {
  if (!drag) return;
  if (drag.type === 'pan') {
    state.panX = drag.panX + (e.clientX - drag.startClientX);
    state.panY = drag.panY + (e.clientY - drag.startClientY);
    drawScene();
    return;
  }
  const dxWorld = (e.clientX - drag.startClientX) / lastScale;
  const dyWorld = -(e.clientY - drag.startClientY) / lastScale;
  if (drag.type === 'die') {
    const die = state.dies[drag.index];
    die.posX = drag.startPosX + dxWorld;
    die.posY = drag.startPosY + dyWorld;
    updateSipInputs();
    drawScene();
  } else if (drag.type === 'sip') {
    state.sip.posX = drag.startPosX + dxWorld;
    state.sip.posY = drag.startPosY + dyWorld;
    state.sip.auto = false;
    drawScene();
  }
});

window.addEventListener('mouseup', (e) => {
  // click（移動 < 5px）→ 選取/取消腳位
  if (pendingClick) {
    const moved = Math.hypot(e.clientX - pendingClick.clientX, e.clientY - pendingClick.clientY);
    if (moved < 5) {
      const hit = findPadAt(pendingClick.worldPoint);
      if (hit) {
        const already = state.selectedPad
          && state.selectedPad.dieIndex === hit.dieIndex
          && String(state.selectedPad.label) === String(hit.pad.label);
        state.selectedPad = already ? null : { dieIndex: hit.dieIndex, label: hit.pad.label };
      } else {
        state.selectedPad = null;
      }
      drawScene();
    }
    pendingClick = null;
  }
  drag = null;
  canvas.style.cursor = 'grab';
});

// Touch 支援（僅畫布平移，簡化版）
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

canvas.addEventListener('wheel', (e) => {
  e.preventDefault();
  const rect = canvas.getBoundingClientRect();
  const cursorCssX = e.clientX - rect.left;
  const cursorCssY = e.clientY - rect.top;
  const factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
  const prevZoom = state.zoomMultiplier;
  const nextZoom = Math.min(Math.max(prevZoom * factor, 0.05), 30);
  state.panX = cursorCssX - (cursorCssX - state.panX) * (nextZoom / prevZoom);
  state.panY = cursorCssY - (cursorCssY - state.panY) * (nextZoom / prevZoom);
  state.zoomMultiplier = nextZoom;
  drawScene();
}, { passive: false });

window.addEventListener('resize', drawScene);
drawScene();
