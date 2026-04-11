"""
HHPDI API — Job Manager
内存级任务追踪，支持 pending / running / done / failed / cancelled 五种状态
"""
from __future__ import annotations

import threading
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Optional


class Job:
    """单个异步任务的状态容器"""

    def __init__(self, job_id: str, job_type: str):
        self.job_id = job_id
        # parse | convert | annotate | pipeline
        self.job_type = job_type
        self.status = "pending"
        self.progress: Dict[str, Any] = {
            "current": 0, "total": 1, "message": "等待中..."
        }
        self.result: Optional[Dict[str, Any]] = None
        self.error: Optional[str] = None
        self.created_at = datetime.now(timezone.utc).isoformat()
        # 外部调用 cancel_event.set() 通知工作线程停止
        self.cancel_event = threading.Event()
        self._lock = threading.Lock()

    # ── 状态更新（线程安全）──────────────────────────────────

    def update_progress(self, current: int, total: int, message: str) -> None:
        with self._lock:
            self.progress = {
                "current": current, "total": total, "message": message
            }

    def set_running(self) -> None:
        with self._lock:
            self.status = "running"

    def set_done(self, result: Dict[str, Any]) -> None:
        with self._lock:
            self.status = "done"
            self.result = result

    def set_failed(self, error: str) -> None:
        with self._lock:
            self.status = "failed"
            self.error = error

    def cancel(self) -> None:
        self.cancel_event.set()
        with self._lock:
            if self.status in ("pending", "running"):
                self.status = "cancelled"

    # ── 序列化 ───────────────────────────────────────────────

    def to_dict(self) -> Dict[str, Any]:
        with self._lock:
            return {
                "job_id":     self.job_id,
                "job_type":   self.job_type,
                "status":     self.status,
                "progress":   dict(self.progress),
                "result":     self.result,
                "error":      self.error,
                "created_at": self.created_at,
            }


class JobManager:
    """全局任务注册表（进程内单例）"""

    def __init__(self):
        self._jobs: Dict[str, Job] = {}
        self._lock = threading.Lock()

    def create_job(self, job_type: str) -> Job:
        job = Job(str(uuid.uuid4()), job_type)
        with self._lock:
            self._jobs[job.job_id] = job
        return job

    def get_job(self, job_id: str) -> Optional[Job]:
        with self._lock:
            return self._jobs.get(job_id)

    def list_jobs(self) -> list:
        with self._lock:
            return [j.to_dict() for j in self._jobs.values()]

    def delete_job(self, job_id: str) -> bool:
        with self._lock:
            if job_id in self._jobs:
                del self._jobs[job_id]
                return True
            return False


# 全局单例
job_manager = JobManager()
