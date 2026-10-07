#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HTML 看板生成器和仅监听本机的自动刷新服务。"""

from __future__ import annotations

import json
import logging
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from snapshot_store import snapshot_rows, write_json_atomic, write_text_atomic


BASE_DIR = Path(__file__).resolve().parent
FALLBACK_TEMPLATE = """<!doctype html>
<html lang="zh-CN"><meta charset="utf-8"><title>B站热门视频</title>
<body><pre id="fallback">__DASHBOARD_DATA__</pre></body></html>"""


class DashboardRuntime:
    def __init__(
        self,
        html_file: Path,
        data_file: Path,
        shortcut_file: Path,
        refresh_seconds: int,
        template_file: Path | None = None,
    ) -> None:
        self.html_file = html_file
        self.data_file = data_file
        self.shortcut_file = shortcut_file
        self.refresh_seconds = max(1, refresh_seconds)
        self.template_file = template_file or (BASE_DIR / "dashboard_template.html")
        self._lock = threading.RLock()

    def update(
        self,
        snapshot: dict[str, Any],
        excel_status: str,
        pending_count: int,
        excel_message: str = "",
    ) -> dict[str, Any]:
        rows = snapshot_rows(snapshot)
        videos: list[dict[str, Any]] = []
        for row in rows:
            videos.append(
                {
                    "rank": int(row.get("rank") or 0),
                    "rank_change": row.get("rank_change"),
                    "previous_rank": row.get("previous_rank"),
                    "cover": str(row.get("cover") or ""),
                    "title": str(row.get("title") or ""),
                    "partition": str(row.get("partition") or "未分区"),
                    "view": int(row.get("view") or 0),
                    "like": int(row.get("like") or 0),
                    "coin": int(row.get("coin") or 0),
                    "favorite": int(row.get("favorite") or 0),
                    "share": int(row.get("share") or 0),
                    "bvid": str(row.get("bvid") or ""),
                    "owner": str(row.get("owner") or ""),
                    "link": str(row.get("link") or ""),
                }
            )

        payload = {
            "generated_at": datetime.now()
            .astimezone()
            .replace(tzinfo=None)
            .isoformat(timespec="seconds"),
            "captured_at": str(snapshot.get("captured_at") or ""),
            "snapshot_id": str(snapshot.get("snapshot_id") or ""),
            "source": "哔哩哔哩全站热门排行",
            "excel_status": excel_status,
            "excel_message": excel_message,
            "pending_count": max(0, int(pending_count)),
            "video_count": len(videos),
            "videos": videos,
        }

        with self._lock:
            write_json_atomic(self.data_file, payload)
            template = self._read_template()
            data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
            html_text = (
                template.replace("__DASHBOARD_DATA__", data_json)
                .replace("__REFRESH_SECONDS__", str(self.refresh_seconds))
                .replace("__GENERATED_AT__", payload["generated_at"])
            )
            write_text_atomic(self.html_file, html_text)
        return payload

    def _read_template(self) -> str:
        if self.template_file.exists():
            return self.template_file.read_text(encoding="utf-8")
        return FALLBACK_TEMPLATE

    def write_shortcut(self, url: str) -> None:
        content = f"[InternetShortcut]\r\nURL={url}\r\n"
        write_text_atomic(self.shortcut_file, content)


class DashboardHTTPServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(
        self,
        server_address: tuple[str, int],
        html_file: Path,
        data_file: Path,
        refresh_event: threading.Event | None = None,
    ) -> None:
        self.html_file = html_file
        self.data_file = data_file
        self.refresh_event = refresh_event or threading.Event()
        super().__init__(server_address, DashboardRequestHandler)


class DashboardRequestHandler(BaseHTTPRequestHandler):
    server_version = "BilibiliHotDashboard/1.0"

    def do_GET(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler API
        path = urlsplit(self.path).path
        if path in {"", "/", "/index.html", "/dashboard"}:
            self._send_file(self.server.html_file, "text/html; charset=utf-8")
            return
        if path == "/api/data":
            self._send_file(self.server.data_file, "application/json; charset=utf-8")
            return
        if path == "/health":
            self._send_json({"ok": True})
            return
        self.send_error(404, "Not Found")

    def _send_file(self, file_path: Path, content_type: str) -> None:
        if not file_path.exists():
            self.send_error(503, "Dashboard data is not ready")
            return
        try:
            content = file_path.read_bytes()
        except OSError as exc:
            self.send_error(500, f"Unable to read dashboard: {exc}")
            return
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self) -> None:  # noqa: N802 - BaseHTTPRequestHandler API
        path = urlsplit(self.path).path
        if path == "/api/refresh":
            self.server.refresh_event.set()
            self._send_json({"ok": True, "message": "刷新请求已提交"})
            return
        self.send_error(404, "Not Found")

    def _send_json(self, payload: dict[str, Any]) -> None:
        content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format: str, *args: Any) -> None:
        logging.debug("Dashboard HTTP: " + format, *args)


def start_dashboard_server(
    host: str,
    port: int,
    html_file: Path,
    data_file: Path,
    refresh_event: threading.Event | None = None,
) -> tuple[DashboardHTTPServer, str]:
    last_error: OSError | None = None
    server: DashboardHTTPServer | None = None
    actual_port = port

    for candidate in range(port, port + 20):
        try:
            server = DashboardHTTPServer(
                (host, candidate),
                html_file,
                data_file,
                refresh_event=refresh_event,
            )
            actual_port = int(server.server_address[1])
            break
        except OSError as exc:
            last_error = exc

    if server is None:
        raise RuntimeError(f"无法启动本地看板服务：{last_error}") from last_error

    thread = threading.Thread(
        target=server.serve_forever,
        name="bilibili-dashboard-server",
        daemon=True,
    )
    thread.start()

    browser_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
    url = f"http://{browser_host}:{actual_port}/"
    return server, url