from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse
import json
import os

from openpyxl import load_workbook

HOST = "127.0.0.1"
PORT = 8000


def strip_extension(file_name: str) -> str:
    return Path(file_name).stem


def normalize_die_name(file_name: str) -> str:
    stem = strip_extension(file_name).strip()
    if stem.lower().endswith("die"):
        stem = stem[:-3]
    return stem or "UnnamedDie"


def round3(value: float) -> float:
    return round(float(value), 3)


def infer_die_size_from_pads(pads: list[dict]) -> dict:
    if not pads:
        return {"x": 1000.0, "y": 1000.0}

    max_x = max(pad["x"] + pad["width"] / 2 for pad in pads)
    max_y = max(pad["y"] + pad["height"] / 2 for pad in pads)
    return {"x": round3(max(max_x, 1.0)), "y": round3(max(max_y, 1.0))}


def build_die_asset(die_name: str, die_size_x: float, die_size_y: float, pads: list[dict]) -> dict:
    pad_items = []
    for pad in pads:
        item = {
            "label": str(pad["label"]),
            "name": str(pad.get("name", "")),
            "x": round3(pad["x"]),
            "y": round3(pad["y"]),
            "width": round3(pad["width"]),
            "height": round3(pad["height"]),
        }
        if pad.get("wirebond") is not None and str(pad["wirebond"]).strip() != "":
            wb = pad["wirebond"]
            if isinstance(wb, float) and wb.is_integer():
                wb = int(wb)
            item["wirebond"] = str(wb).strip()
        pad_items.append(item)
    return {
        "format_version": "3.0",
        "unit": "um",
        "die": {
            "diename": die_name,
            "diesize": {
                "x": round3(die_size_x),
                "y": round3(die_size_y),
            },
        },
        "pads": pad_items,
    }


def value_to_key(value) -> str:
    return (
        str(value or "")
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("(", "")
        .replace(")", "")
        .replace("（", "")
        .replace("）", "")
        .replace("-", "")
    )


def parse_number(value, field_name: str) -> float:
    try:
        return float(value)
    except Exception as exc:
        raise ValueError(f"欄位 {field_name} 含有無法解析的數值：{value}") from exc


def parse_rows_as_standard(rows: list[list], fallback_name: str):
    header_row_index = -1
    for idx, row in enumerate(rows):
        if any(value_to_key(cell) == "diename" for cell in row):
            header_row_index = idx
            break

    if header_row_index == -1:
        return None

    header_row = rows[header_row_index]
    index_by_key = {value_to_key(cell): index for index, cell in enumerate(header_row)}
    required = ["label", "x", "y", "width", "height"]
    for key in required:
        if key not in index_by_key:
            raise ValueError(f"缺少必要欄位：{key}")

    pads = []
    die_name = ""
    die_size_x = None
    die_size_y = None

    for row in rows[header_row_index + 1 :]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        label = row[index_by_key["label"]] if index_by_key["label"] < len(row) else None
        if label is None or str(label).strip() == "":
            continue

        current_die_name = row[index_by_key["diename"]] if "diename" in index_by_key and index_by_key["diename"] < len(row) else die_name or fallback_name
        current_size_x = row[index_by_key["diesizex"]] if "diesizex" in index_by_key and index_by_key["diesizex"] < len(row) else None
        current_size_y = row[index_by_key["diesizey"]] if "diesizey" in index_by_key and index_by_key["diesizey"] < len(row) else None

        die_name = str(current_die_name or fallback_name).strip()
        if current_size_x is not None and current_size_y is not None:
            die_size_x = parse_number(current_size_x, "diesize_x")
            die_size_y = parse_number(current_size_y, "diesize_y")

        pads.append(
            {
                "label": label,
                "name": row[index_by_key["name"]] if "name" in index_by_key and index_by_key["name"] < len(row) else "",
                "x": parse_number(row[index_by_key["x"]], "x"),
                "y": parse_number(row[index_by_key["y"]], "y"),
                "width": parse_number(row[index_by_key["width"]], "width"),
                "height": parse_number(row[index_by_key["height"]], "height"),
            }
        )

    if not pads:
        raise ValueError("Excel 中找不到有效 Pad 資料。")

    inferred_size = infer_die_size_from_pads(pads)
    return {
        "dieAsset": build_die_asset(die_name or fallback_name, die_size_x or inferred_size["x"], die_size_y or inferred_size["y"], pads),
        "warning": "Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。" if die_size_x is None or die_size_y is None else "",
        "parser": "standard",
    }


def parse_rows_as_grouped_sample(rows: list[list], fallback_name: str):
    header = rows[0] if rows else []
    signature = "|".join(
        [
            value_to_key(header[0] if len(header) > 0 else None),
            value_to_key(header[1] if len(header) > 1 else None),
            value_to_key(header[2] if len(header) > 2 else None),
            value_to_key(header[4] if len(header) > 4 else None),
        ]
    )

    if signature != "padno|name|coordinatexy|padsizexy":
        return None

    pads = []
    for row in rows[2:]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        label = row[0] if len(row) > 0 else None
        if label is None or str(label).strip() == "":
            continue
        pads.append(
            {
                "label": label,
                "name": row[1] if len(row) > 1 else "",
                "x": parse_number(row[2], "Coordinate(XY).X"),
                "y": parse_number(row[3], "Coordinate(XY).Y"),
                "width": parse_number(row[4], "PadSize(XY).X"),
                "height": parse_number(row[5], "PadSize(XY).Y"),
            }
        )

    if not pads:
        raise ValueError("範例 Excel 中找不到有效 Pad 資料。")

    inferred_size = infer_die_size_from_pads(pads)
    return {
        "dieAsset": build_die_asset(fallback_name, inferred_size["x"], inferred_size["y"], pads),
        "warning": "此 Excel 未包含 diename / diesize_x / diesize_y，已改用檔名與 Pad 外接範圍推估。",
        "parser": "grouped-sample",
    }


def parse_rows_as_meta_header(rows: list[list], fallback_name: str):
    """Handles format where first rows contain die meta key-value pairs,
    followed by a pad header row then pad data.

    Expected layout:
      Row 0: ['diename', <value>, ...]          (optional row)
      Row 1: ['diesize_x', <val>, 'diesize_y', <val>, ...]
      Row N: ['label', 'name', 'x', 'y', 'width', 'height', ...]
      Row N+1...: pad data
    """
    # Collect key-value meta rows at the top (rows where col-A is a known meta key)
    meta_keys = {"diename", "diesize_x", "diesize_y", "diesizex", "diesizey"}
    meta: dict = {}
    pad_header_index = -1

    for idx, row in enumerate(rows):
        key0 = value_to_key(row[0] if row else None)
        if key0 in meta_keys:
            # Scan all (key, value) pairs in the row
            i = 0
            while i < len(row) - 1:
                k = value_to_key(row[i])
                v = row[i + 1]
                if k in meta_keys and v is not None:
                    meta[k] = v
                i += 2
        elif key0 == "label":
            pad_header_index = idx
            break

    if pad_header_index == -1 or not meta:
        return None

    # Require at least diename OR at least one diesize to treat this as meta-header
    if "diename" not in meta and "diesizex" not in meta and "diesize_x" not in meta:
        return None

    die_name = str(meta.get("diename", fallback_name)).strip() or fallback_name
    size_x_raw = meta.get("diesizex") or meta.get("diesize_x")
    size_y_raw = meta.get("diesizey") or meta.get("diesize_y")
    die_size_x = parse_number(size_x_raw, "diesize_x") if size_x_raw is not None else None
    die_size_y = parse_number(size_y_raw, "diesize_y") if size_y_raw is not None else None

    header_row = rows[pad_header_index]
    index_by_key = {value_to_key(cell): index for index, cell in enumerate(header_row)}

    required = ["label", "x", "y", "width", "height"]
    for key in required:
        if key not in index_by_key:
            raise ValueError(f"缺少必要 Pad 欄位：{key}")

    pads = []
    for row in rows[pad_header_index + 1:]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row):
            continue
        label = row[index_by_key["label"]] if index_by_key["label"] < len(row) else None
        if label is None or str(label).strip() == "":
            continue
        pad = {
            "label": label,
            "name": row[index_by_key["name"]] if "name" in index_by_key and index_by_key["name"] < len(row) else "",
            "x": parse_number(row[index_by_key["x"]], "x"),
            "y": parse_number(row[index_by_key["y"]], "y"),
            "width": parse_number(row[index_by_key["width"]], "width"),
            "height": parse_number(row[index_by_key["height"]], "height"),
        }
        if "wirebond" in index_by_key and index_by_key["wirebond"] < len(row):
            pad["wirebond"] = row[index_by_key["wirebond"]]
        pads.append(pad)

    if not pads:
        raise ValueError("meta-header 格式 Excel 中找不到有效 Pad 資料。")

    inferred = infer_die_size_from_pads(pads)
    return {
        "dieAsset": build_die_asset(
            die_name,
            die_size_x if die_size_x is not None else inferred["x"],
            die_size_y if die_size_y is not None else inferred["y"],
            pads,
        ),
        "warning": "Excel 未提供完整 Die 尺寸，已使用 Pad 外接範圍推估。" if die_size_x is None or die_size_y is None else "",
        "parser": "meta-header",
    }


def parse_workbook_bytes(payload: bytes, file_name: str) -> dict:
    workbook = load_workbook(BytesIO(payload), data_only=True)
    fallback_name = normalize_die_name(file_name)

    def parse_sheet(sheet_name: str):
        worksheet = workbook[sheet_name]
        rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
        return (
            parse_rows_as_meta_header(rows, sheet_name)
            or parse_rows_as_standard(rows, sheet_name)
            or parse_rows_as_grouped_sample(rows, sheet_name)
        )

    # 多工作表：每個 sheet 視為一顆晶片，組成 SiP asset
    parsed_sheets = []
    for sheet_name in workbook.sheetnames:
        try:
            result = parse_sheet(sheet_name)
        except ValueError:
            result = None
        if result is not None:
            parsed_sheets.append((sheet_name, result))

    if len(parsed_sheets) >= 2:
        dies = []
        warnings = []
        for sheet_name, result in parsed_sheets:
            asset = result["dieAsset"]
            dies.append({
                "sheet": sheet_name,
                "diename": asset["die"]["diename"],
                "diesize": asset["die"]["diesize"],
                "pads": asset["pads"],
            })
            if result.get("warning"):
                warnings.append(f"{sheet_name}: {result['warning']}")
        return {
            "sipAsset": {
                "format_version": "4.0",
                "unit": "um",
                "dies": dies,
            },
            "warning": "\n".join(warnings),
            "parser": "multi-sheet",
        }

    if len(parsed_sheets) == 1:
        result = parsed_sheets[0][1]
        # 單 sheet 仍用檔名作為 fallback die name
        if result["dieAsset"]["die"]["diename"] in workbook.sheetnames:
            result["dieAsset"]["die"]["diename"] = fallback_name
        return result

    raise ValueError("無法辨識 Excel 欄位格式。支援格式：\n1. meta-header（前幾列為 diename/diesize，再接 label/x/y/width/height 標頭）\n2. 標準格式（單列標頭含 diename 欄）\n3. GBB898die 範例格式（雙列標頭）\n多工作表檔案：每個 sheet 各自符合上述格式即可組成 SiP。")


class DieRequestHandler(SimpleHTTPRequestHandler):
    def _write_json(self, status_code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/sample-die-asset":
            sample_path = Path("GBB898die.xlsx")
            if not sample_path.exists():
                self._write_json(404, {"error": "找不到 GBB898die.xlsx"})
                return
            try:
                result = parse_workbook_bytes(sample_path.read_bytes(), sample_path.name)
                self._write_json(200, result)
            except Exception as exc:
                self._write_json(400, {"error": str(exc)})
            return

        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/parse-excel":
            self._write_json(404, {"error": "Unsupported endpoint"})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            payload = self.rfile.read(content_length)
            file_name = parse_qs(parsed.query).get("filename", ["ImportedDie.xlsx"])[0]
            result = parse_workbook_bytes(payload, file_name)
            self._write_json(200, result)
        except Exception as exc:
            self._write_json(400, {"error": str(exc)})

if __name__ == "__main__":
    root = Path(__file__).resolve().parent
    os.chdir(root)
    print(f"Serving {root} at http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), DieRequestHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()
