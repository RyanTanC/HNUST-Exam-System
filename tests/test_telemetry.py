"""遥测上报可靠性测试."""

import json
import os

from hnust_exam.services import telemetry


class DummyResponse:
    def __init__(self, ok: bool = True) -> None:
        self.ok = ok

    def raise_for_status(self) -> None:
        if not self.ok:
            raise RuntimeError("HTTP error")


def test_post_checks_http_status(monkeypatch):
    """服务端返回 4xx/5xx 时应视为失败，不能静默当作成功."""
    monkeypatch.setattr(
        telemetry.requests,
        "post",
        lambda *args, **kwargs: DummyResponse(ok=False),
    )

    assert telemetry._post("/api/heartbeat", {"device_id": "dev"}) is False


def test_failed_event_stays_in_queue(tmp_path, monkeypatch):
    """网络失败时事件留在本地队列，供下次启动补发."""
    queue_file = os.path.join(tmp_path, "telemetry_queue.json")

    import hnust_exam.utils.constants as const
    monkeypatch.setattr(const, "_CONFIG_DIR", str(tmp_path))
    monkeypatch.setattr(const, "TELEMETRY_QUEUE_FILE", queue_file)
    monkeypatch.setattr(telemetry, "_start_drain_async", lambda: None)
    monkeypatch.setattr(telemetry.requests, "post", lambda *args, **kwargs: DummyResponse(ok=False))

    telemetry.send_heartbeat("dev123")
    telemetry._drain_queue()

    with open(queue_file, "r", encoding="utf-8") as f:
        queue = json.load(f)
    assert len(queue) == 1
    assert queue[0]["path"] == "/api/heartbeat"
    assert queue[0]["attempts"] == 1


def test_successful_event_removed_from_queue(tmp_path, monkeypatch):
    """发送成功后从队列移除，避免重复上报."""
    queue_file = os.path.join(tmp_path, "telemetry_queue.json")

    import hnust_exam.utils.constants as const
    monkeypatch.setattr(const, "_CONFIG_DIR", str(tmp_path))
    monkeypatch.setattr(const, "TELEMETRY_QUEUE_FILE", queue_file)
    monkeypatch.setattr(telemetry, "_start_drain_async", lambda: None)
    monkeypatch.setattr(telemetry.requests, "post", lambda *args, **kwargs: DummyResponse(ok=True))

    telemetry.send_submit_score("dev123", "试卷A", 88.5, 1200, "单选")
    telemetry._drain_queue()

    with open(queue_file, "r", encoding="utf-8") as f:
        queue = json.load(f)
    assert queue == []
