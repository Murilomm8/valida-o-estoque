#!/usr/bin/env python3
import base64
import io
import json
import os
import posixpath
import zipfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def col_to_index(cell_ref: str) -> int:
    col = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return max(idx - 1, 0)


def parse_xlsx_bytes(data: bytes):
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", NS):
                texts = [t.text or "" for t in si.findall(".//main:t", NS)]
                shared_strings.append("".join(texts))

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        first_sheet = wb.find("main:sheets/main:sheet", NS)
        if first_sheet is None:
            return []
        rel_id = first_sheet.attrib.get(f"{{{NS['rel']}}}id")
        if not rel_id:
            return []

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels.findall("pkgrel:Relationship", NS):
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib.get("Target")
                break
        if not target:
            return []

        target = posixpath.normpath(posixpath.join("xl", target))
        sheet_xml = ET.fromstring(zf.read(target))

        rows = []
        for row in sheet_xml.findall("main:sheetData/main:row", NS):
            values = {}
            for cell in row.findall("main:c", NS):
                ref = cell.attrib.get("r", "A1")
                cidx = col_to_index(ref)
                ctype = cell.attrib.get("t")
                value = ""

                if ctype == "s":
                    v = cell.find("main:v", NS)
                    if v is not None and v.text is not None:
                        try:
                            value = shared_strings[int(v.text)]
                        except Exception:
                            value = ""
                elif ctype == "inlineStr":
                    t = cell.find("main:is/main:t", NS)
                    value = t.text if t is not None and t.text is not None else ""
                else:
                    v = cell.find("main:v", NS)
                    value = v.text if v is not None and v.text is not None else ""

                values[cidx] = value

            if values:
                max_idx = max(values)
                row_list = [values.get(i, "") for i in range(max_idx + 1)]
                rows.append(row_list)

        if not rows:
            return []

        headers = [str(h).strip() for h in rows[0]]
        normalized = [h if h else f"COL_{i+1}" for i, h in enumerate(headers)]
        out = []
        for row in rows[1:]:
            obj = {normalized[i]: row[i] if i < len(row) else "" for i in range(len(normalized))}
            out.append(obj)
        return out


class Handler(SimpleHTTPRequestHandler):
    def translate_path(self, path):
        path = path.split("?", 1)[0].split("#", 1)[0]
        path = posixpath.normpath(unquote(path))
        parts = [p for p in path.split("/") if p and p not in (".", "..")]
        full = ROOT
        for part in parts:
            full = full / part
        return str(full)

    def do_POST(self):
        if self.path != "/api/parse-xlsx":
            self.send_error(404, "Not Found")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(content_length)
            payload = json.loads(raw.decode("utf-8"))
            b64 = payload.get("data", "")
            if not b64:
                raise ValueError("Campo data ausente")
            data = base64.b64decode(b64)
            rows = parse_xlsx_bytes(data)
            body = json.dumps({"rows": rows}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        except Exception as exc:
            body = json.dumps({"error": str(exc)}).encode("utf-8")
            self.send_response(400)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "4173"))
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"Servidor em http://0.0.0.0:{port}")
    server.serve_forever()
