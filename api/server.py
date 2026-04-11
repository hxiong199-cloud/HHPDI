"""
HHPDI REST API Server
符合 RESTful API + OpenAPI 规范（FastAPI 自动生成 /docs 文档）

端口: 默认 8765，可通过环境变量 HHPDI_API_PORT 修改
地址: 默认 127.0.0.1，可通过环境变量 HHPDI_API_HOST 修改
"""
from __future__ import annotations

import os
import shutil
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from api.job_manager import Job, job_manager

# ── 工作目录 ─────────────────────────────────────────────────

_API_JOBS_DIR = Path.home() / ".docflow" / "api_jobs"
_API_JOBS_DIR.mkdir(parents=True, exist_ok=True)


# ══════════════════════════════════════════════════════════════
#  FastAPI 应用
# ══════════════════════════════════════════════════════════════

app = FastAPI(
    title="HHPDI 文档处理智能助手 API",
    description="""
## HHPDI Document Processing REST API

将桌面端的四大核心功能以 HTTP 接口形式开放给外部程序调用。

### 异步任务模式

所有处理接口均立即返回 `job_id`（HTTP 202），调用方通过
`GET /api/v1/jobs/{job_id}` 轮询状态，完成后通过
`GET /api/v1/jobs/{job_id}/files/{filename}` 下载结果文件。

### 四大核心功能

| 功能 | 接口 | 说明 |
|------|------|------|
| 文档解析 | `POST /api/v1/jobs/parse` | PDF / Word → Markdown + 图片 |
| MD→Word | `POST /api/v1/jobs/convert` | Markdown → .docx |
| 数据标注 | `POST /api/v1/jobs/annotate` | Markdown → Dify 知识库格式 |
| 全流水线 | `POST /api/v1/jobs/pipeline` | 解析 → 标注 → 转Word |

### 任务状态说明

`pending` → `running` → `done` / `failed` / `cancelled`
""",
    version="1.0.0",
    contact={"name": "HHPDI Team"},
    openapi_tags=[
        {"name": "系统",     "description": "健康检查、版本信息"},
        {"name": "配置",     "description": "读写软件全局配置"},
        {"name": "核心功能", "description": "文档解析 / 转换 / 标注 / 流水线"},
        {"name": "任务管理", "description": "查询、取消、下载任务结果"},
    ],
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ══════════════════════════════════════════════════════════════
#  内部工具函数
# ══════════════════════════════════════════════════════════════

def _job_dir(job_id: str) -> Path:
    d = _API_JOBS_DIR / job_id
    d.mkdir(parents=True, exist_ok=True)
    return d


def _output_dir(job_id: str) -> Path:
    d = _job_dir(job_id) / "output"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _save_upload(upload: UploadFile, dest: Path) -> None:
    with open(dest, "wb") as f:
        shutil.copyfileobj(upload.file, f)


def _get_job_or_404(job_id: str) -> Job:
    job = job_manager.get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"任务 '{job_id}' 不存在")
    return job


def _file_info(path: Path, job_id: str) -> Dict[str, Any]:
    return {
        "name":         path.name,
        "type":         path.suffix.lstrip("."),
        "size":         path.stat().st_size,
        "download_url": f"/api/v1/jobs/{job_id}/files/{path.name}",
    }


def _copy_to_output(src: str, out_dir: Path, job_id: str) -> Optional[Dict]:
    """将解析结果文件复制到任务 output 目录，返回文件信息字典"""
    if not src:
        return None
    p = Path(src)
    if not p.exists():
        return None
    dst = out_dir / p.name
    shutil.copy2(p, dst)
    return _file_info(dst, job_id)


# ══════════════════════════════════════════════════════════════
#  Pydantic 模型
# ══════════════════════════════════════════════════════════════

class ConfigUpdate(BaseModel):
    model_mode:    Optional[str]  = Field(None, description="模型模式: online | local")
    online:        Optional[dict] = Field(None, description="在线模型配置")
    local:         Optional[dict] = Field(None, description="本地模型配置")
    parse_options: Optional[dict] = Field(None, description="解析选项")


# ══════════════════════════════════════════════════════════════
#  系统接口
# ══════════════════════════════════════════════════════════════

@app.get(
    "/api/v1/health",
    tags=["系统"],
    summary="健康检查",
    response_description="服务状态",
)
def health():
    """检查 API 服务是否正常运行"""
    return {"status": "ok", "service": "HHPDI API", "version": "1.0.0"}


# ══════════════════════════════════════════════════════════════
#  配置接口
# ══════════════════════════════════════════════════════════════

@app.get(
    "/api/v1/config",
    tags=["配置"],
    summary="获取当前配置",
)
def get_config_api():
    """
    获取当前软件配置。

    **注意**：API Key 会被脱敏（只显示末4位）。
    """
    from config.settings import get_config
    cfg = get_config()
    # 脱敏
    import copy
    safe = copy.deepcopy(cfg)
    for section in ("online", "local"):
        key = safe.get(section, {}).get("api_key", "")
        if key:
            safe[section]["api_key"] = ("*" * max(0, len(key) - 4)) + key[-4:]
    return safe


@app.put(
    "/api/v1/config",
    tags=["配置"],
    summary="更新配置",
)
def update_config_api(body: ConfigUpdate):
    """
    更新软件配置。只需传入要修改的字段，其余保持不变。

    **示例**（更新 VLM 模型）：
    ```json
    {
      "online": {
        "model": "gpt-4o",
        "api_key": "sk-xxx"
      }
    }
    ```
    """
    from config.settings import update_config
    update_config(body.model_dump(exclude_none=True))
    return {"status": "ok", "message": "配置已更新"}


# ══════════════════════════════════════════════════════════════
#  任务管理接口
# ══════════════════════════════════════════════════════════════

@app.get(
    "/api/v1/jobs",
    tags=["任务管理"],
    summary="列出所有任务",
)
def list_jobs():
    """返回所有已提交任务的状态列表"""
    return job_manager.list_jobs()


@app.get(
    "/api/v1/jobs/{job_id}",
    tags=["任务管理"],
    summary="查询任务状态",
)
def get_job(job_id: str):
    """
    查询任务状态与进度。

    **status 字段说明：**
    - `pending`   — 排队等待
    - `running`   — 执行中，`progress` 字段实时更新
    - `done`      — 已完成，`result.files` 包含可下载文件列表
    - `failed`    — 失败，`error` 字段包含错误信息
    - `cancelled` — 已取消
    """
    return _get_job_or_404(job_id).to_dict()


@app.delete(
    "/api/v1/jobs/{job_id}",
    tags=["任务管理"],
    summary="取消/删除任务",
)
def delete_job(job_id: str):
    """
    取消正在运行的任务，或删除已完成任务的记录（同时清理本地文件）。
    """
    job = _get_job_or_404(job_id)
    job.cancel()
    # 清理工作目录
    work_dir = _API_JOBS_DIR / job_id
    if work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)
    job_manager.delete_job(job_id)
    return {"status": "ok", "message": f"任务 '{job_id}' 已取消/删除"}


@app.get(
    "/api/v1/jobs/{job_id}/files/{filename}",
    tags=["任务管理"],
    summary="下载任务结果文件",
    response_class=FileResponse,
)
def download_file(job_id: str, filename: str):
    """
    下载任务完成后的输出文件。

    可下载的文件名列表在 `GET /api/v1/jobs/{job_id}` 的
    `result.files[*].name` 字段中查看。

    **支持的文件类型：** `.md`, `.docx`, `.zip`
    """
    job = _get_job_or_404(job_id)
    if job.status != "done":
        raise HTTPException(
            status_code=400,
            detail=f"任务尚未完成（当前状态: {job.status}）",
        )
    # 安全校验：防止路径穿越
    out_dir = (_API_JOBS_DIR / job_id / "output").resolve()
    target  = (out_dir / filename).resolve()
    if not str(target).startswith(str(out_dir)):
        raise HTTPException(status_code=400, detail="非法文件路径")
    if not target.exists():
        raise HTTPException(status_code=404, detail=f"文件 '{filename}' 不存在")
    return FileResponse(path=str(target), filename=filename)


# ══════════════════════════════════════════════════════════════
#  核心功能 — 文档解析
# ══════════════════════════════════════════════════════════════

@app.post(
    "/api/v1/jobs/parse",
    tags=["核心功能"],
    summary="文档解析（PDF/Word → Markdown）",
    status_code=202,
)
async def submit_parse(
    file: UploadFile = File(..., description="PDF 或 Word 文档（.pdf / .docx / .doc）"),
):
    """
    上传文档，异步解析为 **Markdown + 图片**。

    - 支持格式：`.pdf`、`.docx`、`.doc`（单文件最大 50 MB）
    - 采用 VLM 识别版面、表格、公式；PyMuPDF 提取文字页文本

    **完成后可下载：**
    - `{文件名}.md` — Markdown 文档
    - `{文件名}.docx` — Word 文档
    - `{文件名}.zip` — 打包（含图片）

    **轮询方式：**
    ```
    GET /api/v1/jobs/{job_id}   →  status=done 后
    GET /api/v1/jobs/{job_id}/files/{filename}
    ```
    """
    job     = job_manager.create_job("parse")
    job_dir = _job_dir(job.job_id)
    out_dir = _output_dir(job.job_id)

    orig_name  = file.filename or "document.pdf"
    input_path = job_dir / orig_name
    _save_upload(file, input_path)

    def _run():
        job.set_running()
        try:
            from core.pipeline import run_pipeline

            def _progress(cur, total, msg):
                job.update_progress(cur, total, msg)

            result = run_pipeline(
                str(input_path),
                progress_cb=_progress,
                cancel_event=job.cancel_event,
            )

            if job.cancel_event.is_set():
                return

            if not result.success:
                job.set_failed(result.error or "解析失败")
                return

            files_info = []
            for src in (result.md_path, result.docx_path, result.zip_path):
                info = _copy_to_output(src, out_dir, job.job_id)
                if info:
                    files_info.append(info)

            job.set_done({
                "files":              files_info,
                "stats":              result.stats,
                "md_content_preview": (result.md_content or "")[:500],
            })
        except Exception as exc:
            import traceback
            job.set_failed(f"{exc}\n{traceback.format_exc()}")

    threading.Thread(target=_run, daemon=True).start()
    return {
        "job_id":  job.job_id,
        "status":  "pending",
        "message": "文档解析任务已提交，请通过 job_id 轮询进度",
        "poll_url": f"/api/v1/jobs/{job.job_id}",
    }


# ══════════════════════════════════════════════════════════════
#  核心功能 — MD → Word 转换
# ══════════════════════════════════════════════════════════════

@app.post(
    "/api/v1/jobs/convert",
    tags=["核心功能"],
    summary="Markdown → Word 转换",
    status_code=202,
)
async def submit_convert(
    file: UploadFile = File(
        ..., description="Markdown 文件（.md）"
    ),
    images_zip: Optional[UploadFile] = File(
        None,
        description="图片文件夹打包（.zip，可选）。"
                    "若不提供，Word 中图片将显示为占位符。",
    ),
):
    """
    上传 Markdown 文件（及可选图片 zip），转换为 **.docx** Word 文档。

    支持的 Markdown 语法：标题、段落、表格（pipe / HTML）、图片引用、
    加粗 / 斜体、`####` 分段符、`@@@tags@@@` 标签行。

    **完成后可下载：**
    - `{文件名}.docx` — Word 文档
    """
    job     = job_manager.create_job("convert")
    job_dir = _job_dir(job.job_id)
    out_dir = _output_dir(job.job_id)

    orig_name  = file.filename or "document.md"
    input_path = job_dir / orig_name
    _save_upload(file, input_path)

    images_dir = job_dir / "images"
    images_dir.mkdir(exist_ok=True)
    if images_zip:
        zip_path = job_dir / "images.zip"
        _save_upload(images_zip, zip_path)
        shutil.unpack_archive(str(zip_path), str(images_dir))

    def _run():
        job.set_running()
        try:
            # 延迟导入避免 tkinter 在 import 阶段初始化
            from tools.tool2_converter import run_conversion  # noqa

            stem        = Path(orig_name).stem
            output_path = str(out_dir / f"{stem}.docx")

            job.update_progress(0, 1, "MD → Word 转换中...")
            run_conversion(str(input_path), str(images_dir), output_path)
            job.update_progress(1, 1, "完成")

            files_info = []
            p = Path(output_path)
            if p.exists():
                files_info.append(_file_info(p, job.job_id))

            job.set_done({"files": files_info})
        except Exception as exc:
            import traceback
            job.set_failed(f"{exc}\n{traceback.format_exc()}")

    threading.Thread(target=_run, daemon=True).start()
    return {
        "job_id":   job.job_id,
        "status":   "pending",
        "message":  "MD→Word 转换任务已提交",
        "poll_url": f"/api/v1/jobs/{job.job_id}",
    }


# ══════════════════════════════════════════════════════════════
#  核心功能 — 数据标注
# ══════════════════════════════════════════════════════════════

@app.post(
    "/api/v1/jobs/annotate",
    tags=["核心功能"],
    summary="数据标注（Markdown → Dify 知识库格式）",
    status_code=202,
)
async def submit_annotate(
    file: UploadFile = File(..., description="Markdown 文件（.md）"),
    llm_url: str = Form(
        default="https://api.siliconflow.cn/v1/chat/completions",
        description="LLM 补全接口地址（OpenAI 兼容格式）",
    ),
    llm_key: str = Form(..., description="LLM API Key"),
    llm_model: str = Form(
        default="Pro/deepseek-ai/DeepSeek-V3",
        description="LLM 模型名称",
    ),
    concurrency: int = Form(
        default=3, ge=1, le=10,
        description="并发标注任务数（1–10）",
    ),
):
    """
    上传 Markdown 文件，使用 LLM 进行智能数据标注，
    输出 **Dify 知识库**标准分段格式。

    **输出格式：**
    ```
    ####
    段落内容
    tags: @@@标签1@@@标签2@@@标签3@@@
    ####
    Q:问题
    A:答案
    tags: @@@标签@@@
    ####
    ```

    Dify 导入时将**分段标识符**设为 `####` 即可精确切分。

    **完成后可下载：**
    - `{文件名}_annotated.md` — 带标注的 Markdown
    """
    job     = job_manager.create_job("annotate")
    job_dir = _job_dir(job.job_id)
    out_dir = _output_dir(job.job_id)

    orig_name  = file.filename or "document.md"
    input_path = job_dir / orig_name
    _save_upload(file, input_path)

    def _run():
        job.set_running()
        try:
            from api.annotator_core import annotate_md

            def _progress(cur, total, msg):
                job.update_progress(cur, total, msg)

            out_path = annotate_md(
                str(input_path),
                llm_url=llm_url,
                llm_key=llm_key,
                llm_model=llm_model,
                concurrency=concurrency,
                progress_cb=_progress,
                cancel_event=job.cancel_event,
            )

            dst = out_dir / Path(out_path).name
            shutil.copy2(out_path, dst)

            chunk_count = dst.read_text(encoding='utf-8').count('####')
            job.set_done({
                "files":       [_file_info(dst, job.job_id)],
                "chunk_count": chunk_count,
            })
        except Exception as exc:
            import traceback
            job.set_failed(f"{exc}\n{traceback.format_exc()}")

    threading.Thread(target=_run, daemon=True).start()
    return {
        "job_id":   job.job_id,
        "status":   "pending",
        "message":  "数据标注任务已提交",
        "poll_url": f"/api/v1/jobs/{job.job_id}",
    }


# ══════════════════════════════════════════════════════════════
#  核心功能 — 全流水线
# ══════════════════════════════════════════════════════════════

@app.post(
    "/api/v1/jobs/pipeline",
    tags=["核心功能"],
    summary="全流水线（解析 → 标注 → 转Word）",
    status_code=202,
)
async def submit_pipeline(
    file: UploadFile = File(..., description="PDF 或 Word 文档"),
    llm_url: str = Form(
        default="https://api.siliconflow.cn/v1/chat/completions",
        description="标注用 LLM 接口地址",
    ),
    llm_key: str = Form(..., description="标注用 LLM API Key"),
    llm_model: str = Form(
        default="Pro/deepseek-ai/DeepSeek-V3",
        description="标注用 LLM 模型",
    ),
    concurrency: int = Form(default=3, ge=1, le=10, description="标注并发数"),
    skip_annotate: bool = Form(default=False, description="跳过数据标注步骤"),
    skip_convert: bool = Form(default=False, description="跳过 Word 转换步骤"),
):
    """
    上传文档，执行完整的三阶段流水线：

    1. **文档解析** — PDF/Word → Markdown + 图片（VLM 驱动）
    2. **数据标注** — Markdown → Dify 格式（可通过 `skip_annotate=true` 跳过）
    3. **Word 转换** — 标注后的 Markdown → .docx（可通过 `skip_convert=true` 跳过）

    任意步骤失败不会中断流水线，失败信息记录在 `result.errors` 中。

    **完成后可下载所有输出文件。**
    """
    job     = job_manager.create_job("pipeline")
    job_dir = _job_dir(job.job_id)
    out_dir = _output_dir(job.job_id)

    orig_name  = file.filename or "document.pdf"
    input_path = job_dir / orig_name
    _save_upload(file, input_path)

    def _run():
        job.set_running()
        files_info: List[Dict] = []
        errors: List[str]      = []

        try:
            # ── Step 1：文档解析 ─────────────────────────
            from core.pipeline import run_pipeline

            job.update_progress(0, 3, "[1/3] 文档解析中...")

            def _parse_progress(cur, total, msg):
                job.update_progress(0, 3, f"[1/3] {msg}")

            parse_result = run_pipeline(
                str(input_path),
                progress_cb=_parse_progress,
                cancel_event=job.cancel_event,
            )

            if job.cancel_event.is_set():
                return

            if not parse_result.success:
                job.set_failed(f"解析失败: {parse_result.error}")
                return

            # 复制解析产物
            for src in (parse_result.md_path,
                        parse_result.docx_path,
                        parse_result.zip_path):
                info = _copy_to_output(src, out_dir, job.job_id)
                if info:
                    files_info.append(info)

            md_path    = parse_result.md_path
            images_dir = parse_result.images_dir

            # ── Step 2：数据标注 ─────────────────────────
            annotated_md = md_path
            if not skip_annotate and llm_key:
                job.update_progress(1, 3, "[2/3] 数据标注中...")
                try:
                    from api.annotator_core import annotate_md

                    def _ann_progress(cur, total, msg):
                        job.update_progress(1, 3, f"[2/3] {msg}")

                    annotated_md = annotate_md(
                        md_path,
                        llm_url=llm_url,
                        llm_key=llm_key,
                        llm_model=llm_model,
                        concurrency=concurrency,
                        progress_cb=_ann_progress,
                        cancel_event=job.cancel_event,
                    )
                    info = _copy_to_output(annotated_md, out_dir, job.job_id)
                    if info:
                        files_info.append(info)
                except Exception as exc:
                    errors.append(f"标注失败（已跳过）: {exc}")
                    annotated_md = md_path  # 降级：用原始 MD 继续

                if job.cancel_event.is_set():
                    return

            # ── Step 3：Word 转换 ─────────────────────────
            if not skip_convert:
                job.update_progress(2, 3, "[3/3] 转换 Word 中...")
                try:
                    from tools.tool2_converter import run_conversion  # noqa

                    stem     = Path(annotated_md).stem
                    word_out = out_dir / f"{stem}.docx"
                    run_conversion(annotated_md, images_dir, str(word_out))
                    if word_out.exists():
                        files_info.append(_file_info(word_out, job.job_id))
                except Exception as exc:
                    errors.append(f"Word转换失败（已跳过）: {exc}")

            job.update_progress(3, 3, "全部完成")
            job.set_done({
                "files":  files_info,
                "stats":  parse_result.stats,
                "errors": errors,
            })

        except Exception as exc:
            import traceback
            job.set_failed(f"{exc}\n{traceback.format_exc()}")

    threading.Thread(target=_run, daemon=True).start()
    return {
        "job_id":   job.job_id,
        "status":   "pending",
        "message":  "流水线任务已提交",
        "poll_url": f"/api/v1/jobs/{job.job_id}",
    }
