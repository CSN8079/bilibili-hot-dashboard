#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓取快照的持久化队列，保证 Excel 写入失败时不丢失数据。"""

from __future__ import annotations

import json
import logging
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any


class PendingQueueError(RuntimeError):
    """待补写队列损坏。"""


def _jsonable_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        item = dict(row)
        captured_at = item.get("captured_at")
        if isinstance(captured_at, datetime):
            item["captured_at"] = captured_at.isoformat(timespec="microseconds")
        output.append(item)
    return output


def build_snapshot(rows: list[dict[str, Any]]) -> dict[str, Any]:
    if not rows:
        raise ValueError("不能创建空的数据快照。")

    captured_at = rows[0].get("captured_at")
    if isinstance(captured_at, datetime):
        captured_iso = captured_at.isoformat(timespec="microseconds")
    else:
        captured_iso = datetime.now().astimezone().replace(tzinfo=None).isoformat(
            timespec="microseconds"
        )

    return {
        "version": 1,
        "snapshot_id": (
            datetime.fromisoformat(captured_iso).strftime("%Y%m%dT%H%M%S%f")
            + "-"
            + uuid.uuid4().hex[:8]
        ),
        "captured_at": captured_iso,
        "rows": _jsonable_rows(rows),
    }


def snapshot_rows(snapshot: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source in snapshot.get("rows", []):
        if not isinstance(source, dict):
            continue
        item = dict(source)
        captured_at = item.get("captured_at")
        if isinstance(captured_at, str):
            try:
                item["captured_at"] = datetime.fromisoformat(captured_at)
            except ValueError:
                item["captured_at"] = datetime.now().astimezone().replace(tzinfo=None)
        rows.append(item)
    return rows


def read_pending_snapshots(pending_file: Path) -> list[dict[str, Any]]:
    if not pending_file.exists():
        return []

    snapshots: list[dict[str, Any]] = []
    with pending_file.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            text = line.strip()
            if not text:
                continue
            try:
                snapshot = json.loads(text)
            except json.JSONDecodeError as exc:
                raise PendingQueueError(
                    f"待补写文件第 {line_number} 行损坏：{pending_file}"
                ) from exc
            if not isinstance(snapshot, dict) or not isinstance(snapshot.get("rows"), list):
                raise PendingQueueError(
                    f"待补写文件第 {line_number} 行格式不正确：{pending_file}"
                )
            snapshots.append(snapshot)
    return snapshots


def append_pending_snapshot(
    pending_file: Path,
    snapshot: dict[str, Any],
) -> list[dict[str, Any]]:
    pending_file.parent.mkdir(parents=True, exist_ok=True)
    pending = read_pending_snapshots(pending_file)
    pending.append(snapshot)

    payload = json.dumps(snapshot, ensure_ascii=False, separators=(",", ":"))
    with pending_file.open("a", encoding="utf-8", newline="\n") as file:
        file.write(payload)
        file.write("\n")
        file.flush()
        os.fsync(file.fileno())

    return pending


def clear_pending_snapshots(pending_file: Path) -> None:
    pending_file.unlink(missing_ok=True)


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_file = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    try:
        with temp_file.open("w", encoding="utf-8", newline="\n") as file:
            file.write(text)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temp_file, path)
    finally:
        temp_file.unlink(missing_ok=True)


def write_text_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_file = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        with temp_file.open("w", encoding="utf-8", newline="\n") as file:
            file.write(text)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temp_file, path)
    finally:
        temp_file.unlink(missing_ok=True)