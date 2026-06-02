"""匿名使用统计上报服务.

完全静默：先把事件保存到本地队列，再由后台线程发送；失败会保留到下次启动补发。
"""

from __future__ import annotations

import json
import os
import platform
import threading
import time
import uuid

import requests

from hnust_exam.utils import constants

_TELEMETRY_TIMEOUT = 8  # 秒
_MAX_QUEUE_SIZE = 100
_MAX_SEND_PER_DRAIN = 20
_QUEUE_LOCK = threading.RLock()
_DRAINING = False


def _get_base_url() -> str:
    return constants.TELEMETRY_BASE_URL.rstrip("/")


def _get_queue_file() -> str:
    return constants.TELEMETRY_QUEUE_FILE


def _load_queue() -> list[dict]:
    """读取本地待发送队列，文件损坏时从空队列恢复."""
    try:
        queue_file = _get_queue_file()
        if not os.path.exists(queue_file):
            return []
        with open(queue_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_queue(queue: list[dict]) -> None:
    """原子写入本地队列，避免程序中断时留下半截 JSON."""
    try:
        os.makedirs(constants._CONFIG_DIR, exist_ok=True)
        queue_file = _get_queue_file()
        tmp_file = f"{queue_file}.tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(queue[-_MAX_QUEUE_SIZE:], f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, queue_file)
    except Exception:
        pass


def _post(path: str, data: dict) -> bool:
    """在当前线程发送 POST，返回是否成功."""
    try:
        response = requests.post(f"{_get_base_url()}{path}", json=data, timeout=_TELEMETRY_TIMEOUT)
        response.raise_for_status()
        return True
    except Exception:
        return False


def _enqueue(path: str, data: dict) -> None:
    """先落盘再触发后台发送，避免快速退出导致事件丢失."""
    event = {
        "id": uuid.uuid4().hex,
        "path": path,
        "data": data,
        "created_at": int(time.time()),
        "attempts": 0,
        "next_try": 0,
    }
    with _QUEUE_LOCK:
        queue = _load_queue()
        queue.append(event)
        _save_queue(queue)
    _start_drain_async()


def _retry_delay(attempts: int) -> int:
    return min(3600, 30 * (2 ** min(attempts, 6)))


def _drain_queue() -> None:
    """发送本地队列中的到期事件，失败事件保留并延迟重试."""
    global _DRAINING
    with _QUEUE_LOCK:
        if _DRAINING:
            return
        _DRAINING = True

    try:
        for _ in range(_MAX_SEND_PER_DRAIN):
            with _QUEUE_LOCK:
                queue = _load_queue()
                now = time.time()
                index = next(
                    (i for i, event in enumerate(queue)
                     if float(event.get("next_try", 0)) <= now),
                    None,
                )
                if index is None:
                    return
                event = queue[index]

            success = _post(event.get("path", ""), event.get("data", {}))

            with _QUEUE_LOCK:
                queue = _load_queue()
                index = next(
                    (i for i, item in enumerate(queue)
                     if item.get("id") == event.get("id")),
                    None,
                )
                if index is None:
                    continue
                if success:
                    queue.pop(index)
                else:
                    attempts = int(queue[index].get("attempts", 0)) + 1
                    queue[index]["attempts"] = attempts
                    queue[index]["next_try"] = int(time.time()) + _retry_delay(attempts)
                _save_queue(queue)
    finally:
        with _QUEUE_LOCK:
            _DRAINING = False


def _start_drain_async() -> None:
    """后台线程补发队列."""
    threading.Thread(target=_drain_queue, daemon=True).start()


# ── 公开接口 ──────────────────────────────────────────────

def init_telemetry(config_mgr) -> str:
    """初始化 device_id（如不存在则生成），返回 device_id."""
    cfg = config_mgr.load_config()
    device_id = cfg.get("device_id", "")
    if not device_id:
        device_id = uuid.uuid4().hex[:16]
        cfg["device_id"] = device_id
        config_mgr.save_config(cfg)
    _start_drain_async()
    return device_id


def send_heartbeat(device_id: str) -> None:
    """发送心跳."""
    _enqueue("/api/heartbeat", {
        "device_id": device_id,
        "version": constants.CURRENT_VERSION,
        "os": platform.platform(),
    })


def send_submit_score(device_id: str, exam_name: str, score_pct: float,
                      duration_seconds: int | None = None,
                      question_types: str = "") -> None:
    """交卷后上报成绩."""
    _enqueue("/api/submit-score", {
        "device_id": device_id,
        "exam_name": exam_name,
        "score_pct": score_pct,
        "duration_seconds": duration_seconds,
        "question_types": question_types,
        "version": constants.CURRENT_VERSION,
        "os": platform.platform(),
    })
