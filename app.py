# -*- coding: utf-8 -*-
"""行程智能体 Web 服务（标准库实现，无外部依赖，复用 agent_lib）

启动：python app.py
本地默认监听 127.0.0.1:8788；云端读取 PORT 环境变量并绑定 0.0.0.0
"""
import json
import os
import re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

import agent_lib as L

BASE = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(BASE, "index.html")
# 云端部署必须读取 PORT 环境变量并绑定 0.0.0.0
PORT = int(os.environ.get("PORT", 8788))
HOST = os.environ.get("HOST", "0.0.0.0")


def _json_body(handler):
    length = int(handler.headers.get("Content-Length", 0) or 0)
    raw = handler.rfile.read(length) if length else b"{}"
    try:
        return json.loads(raw.decode("utf-8")) if raw.strip() else {}
    except json.JSONDecodeError:
        return {}


def _send_json(handler, obj, code=200):
    body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    handler.send_response(code)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def _send_file(handler, path, ctype, as_attachment=True):
    with open(path, "rb") as f:
        body = f.read()
    handler.send_response(200)
    handler.send_header("Content-Type", ctype)
    handler.send_header("Content-Length", str(len(body)))
    if as_attachment:
        handler.send_header("Content-Disposition",
                            'attachment; filename="%s"' % os.path.basename(path))
    handler.end_headers()
    handler.wfile.write(body)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urlparse(self.path)
        path = u.path
        if path in ("/", "/index.html"):
            with open(INDEX_PATH, encoding="utf-8") as f:
                html = f.read()
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/api/list":
            rows = L.active_rows()
            recs = [{c: r.get(c, "") for c in L.COLUMNS} for r in rows]
            return _send_json(self, {"rows": recs})
        if path == "/api/dashboard":
            return _send_json(self, {"records": L.get_dashboard_data()})
        if path == "/api/export":
            xlsx = L.export_excel()
            return _send_file(self, xlsx,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        u = urlparse(self.path)
        path = u.path
        data = _json_body(self)

        if path == "/api/parse":
            text = data.get("text", "")
            res = L.parse_oral_text(text)
            proj = res["found"].get("项目名称", "")
            history_level = L.compute_level(L.load_rows(), proj) if proj else 1
            # 优先采用文本中明确提到的当前等级；否则给出历史递增建议
            if res.get("level_hint") is not None:
                res["level"] = res["level_hint"]
                res["level_source"] = "text"
            else:
                res["level"] = history_level
                res["level_source"] = "history"
            return _send_json(self, res)

        if path == "/api/add":
            fields = data.get("fields", {})
            override = data.get("level_override")
            override = int(override) if str(override).strip().isdigit() else None
            rec = L.add_trip(fields, level_override=override)
            return _send_json(self, {"record": rec})

        if path == "/api/edit":
            rid = str(data.get("id", ""))
            updates = data.get("updates", {})
            # 若改了项目名称，层级不自动重算，保持用户确认值
            rec = L.edit_trip(rid, updates)
            return _send_json(self, {"record": rec} if rec else {"error": "not found"})

        if path == "/api/delete":
            rid = str(data.get("id", ""))
            L.soft_delete(rid)
            return _send_json(self, {"ok": True})

        if path == "/api/hard_delete":
            rid = str(data.get("id", ""))
            L.hard_delete(rid)
            return _send_json(self, {"ok": True})

        self.send_response(404)
        self.end_headers()


def main():
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    print("行程智能体 Web 服务已启动: http://%s:%d" % (HOST, PORT))
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.shutdown()


if __name__ == "__main__":
    main()
