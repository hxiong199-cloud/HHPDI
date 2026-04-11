# HHPDI 文档数据处理智能助手

集成四个工具的统一平台，覆盖从原始文档到知识库的完整处理链路：

| 工具 | 功能 |
|------|------|
| ① 文档解析 | PDF / Word → Markdown（VLM 智能解析，支持扫描件） |
| ② MD → Word | Markdown + 图片 → .docx 文档 |
| ③ 数据标注 | 文本打标签 + 表格转 QA（知识库构建） |
| ⚡ 流水线 | 一键串联上述三个步骤，支持批量多文件并行 |

以上四个功能同时提供**桌面 GUI** 和 **REST API** 两种调用方式。

---

## 安装

### 桌面端（必需）

```bash
pip install -r requirements.txt
```

### API 服务（可选）

```bash
pip install fastapi "uvicorn[standard]" python-multipart
```

> 不安装 API 依赖时，桌面程序正常运行，API 功能静默跳过，互不影响。

## 运行

```bash
python main.py
```

启动后：
- 桌面 GUI 在主线程运行
- 若已安装 API 依赖，REST API 自动在后台线程启动，**无需任何额外操作**

---

## 使用流程（桌面端）

1. 点击右上角「模型设置」配置 API Key（VLM 模型用于文档解析）
2. 在「文档解析」选择 PDF 或 Word 文件（支持多选或选择文件夹，最多 10 个），点击「开始批量解析」
3. 解析完成后点击「→ 传递」，在「MD→Word」或「数据标注」中直接使用
4. 或使用「流水线」一键完成全流程

---

## REST API

桌面程序启动时自动开启 API 服务，默认监听：

```
http://127.0.0.1:8765
```

### 在线文档

| 地址 | 说明 |
|------|------|
| `http://127.0.0.1:8765/docs` | Swagger UI（可直接在浏览器测试接口） |
| `http://127.0.0.1:8765/redoc` | ReDoc 文档（阅读友好） |
| `http://127.0.0.1:8765/openapi.json` | OpenAPI JSON Schema |

### 端点一览

| 方法 | 路径 | 功能 |
|------|------|------|
| `GET` | `/api/v1/health` | 健康检查 |
| `GET` | `/api/v1/config` | 读取当前配置（Key 自动脱敏） |
| `PUT` | `/api/v1/config` | 更新配置 |
| `POST` | `/api/v1/jobs/parse` | **文档解析**：上传 PDF/Word，返回 MD + 图片 + docx |
| `POST` | `/api/v1/jobs/convert` | **MD→Word**：上传 MD（及可选图片 zip），返回 docx |
| `POST` | `/api/v1/jobs/annotate` | **数据标注**：上传 MD，返回 Dify 格式 `_annotated.md` |
| `POST` | `/api/v1/jobs/pipeline` | **全流水线**：解析 → 标注 → 转 Word，一次提交 |
| `GET` | `/api/v1/jobs` | 列出所有任务 |
| `GET` | `/api/v1/jobs/{job_id}` | 查询任务状态与进度 |
| `DELETE` | `/api/v1/jobs/{job_id}` | 取消/删除任务及本地文件 |
| `GET` | `/api/v1/jobs/{job_id}/files/{filename}` | 下载任务输出文件 |

### 异步任务模式

所有处理接口（parse / convert / annotate / pipeline）均立即返回 `job_id`（HTTP 202），调用方轮询状态后下载结果：

```
# 1. 提交任务
POST /api/v1/jobs/parse
Content-Type: multipart/form-data
Body: file=<your_document.pdf>

# 返回：
{ "job_id": "xxxxxxxx-...", "status": "pending", "poll_url": "/api/v1/jobs/xxxxxxxx-..." }

# 2. 轮询状态（建议间隔 2–5 秒）
GET /api/v1/jobs/{job_id}

# 返回（执行中）：
{ "status": "running", "progress": { "current": 3, "total": 10, "message": "VLM 分析第 3/10 页..." } }

# 返回（完成）：
{ "status": "done", "result": { "files": [ { "name": "report.md", "download_url": "/api/v1/jobs/.../files/report.md" }, ... ] } }

# 3. 下载结果
GET /api/v1/jobs/{job_id}/files/report.md
```

**任务状态说明：**

| 状态 | 含义 |
|------|------|
| `pending` | 排队等待 |
| `running` | 执行中，`progress` 字段实时更新 |
| `done` | 完成，`result.files` 包含可下载文件列表 |
| `failed` | 失败，`error` 字段包含错误详情 |
| `cancelled` | 已取消 |

### 调用示例（Python）

```python
import requests, time

BASE = "http://127.0.0.1:8765"

# 提交全流水线任务
with open("report.pdf", "rb") as f:
    resp = requests.post(f"{BASE}/api/v1/jobs/pipeline", files={"file": f}, data={
        "llm_key": "sk-xxx",
        "llm_model": "Pro/deepseek-ai/DeepSeek-V3",
    })
job_id = resp.json()["job_id"]

# 轮询等待完成
while True:
    status = requests.get(f"{BASE}/api/v1/jobs/{job_id}").json()
    print(status["progress"]["message"])
    if status["status"] in ("done", "failed", "cancelled"):
        break
    time.sleep(3)

# 下载所有结果文件
for file_info in status["result"]["files"]:
    data = requests.get(f"{BASE}{file_info['download_url']}").content
    with open(file_info["name"], "wb") as f:
        f.write(data)
    print(f"已保存：{file_info['name']}")
```

### 环境变量配置

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `HHPDI_API_HOST` | `127.0.0.1` | 监听地址（设为 `0.0.0.0` 可对局域网开放） |
| `HHPDI_API_PORT` | `8765` | 监听端口 |

```bash
HHPDI_API_PORT=9000 python main.py
```

---

## 导入 Dify 知识库

「数据标注」工具的输出文件（`*_annotated.md`）专为 Dify 知识库格式设计，导入时需按以下说明配置分段参数。

### 输出文件格式说明

标注完成后，每个知识片段的结构如下：

```
正文内容（或表格 QA 问答对）
@@@标签1,标签2,标签3@@@

####
```

- **`@@@...@@@`**：标签标志符，包裹该片段的分类标签
- **`####`**：分段符，用于分隔相邻的知识片段

### Dify 知识库导入配置

在 Dify「知识库 → 创建知识库 → 选择文件」上传 `_annotated.md` 后，**分段设置**按如下填写：

#### 方式一：父子分段（推荐）

适合内容较长、需要层级检索的场景。

| 设置项 | 填写值 | 说明 |
|--------|--------|------|
| 分段模式 | **父子分段** | — |
| 父分段标识符 | `####` | 与工具输出的分段符一致 |
| 子分段标识符 | 换行符（`\n`） | 在父段内按自然段再细分 |
| 最大长度 | 根据模型上下文调整，建议 **500～2000** tokens | — |
| 标签字段 | `@@@` | 告知 Dify 标签的包裹符号 |

> **效果**：父段 = 一个完整知识片段，子段 = 片段内各自然段，检索时优先命中子段、召回完整父段。

#### 方式二：普通分段

适合内容较短、结构扁平的场景。

| 设置项 | 填写值 |
|--------|--------|
| 分段模式 | **自定义** |
| 分段标识符 | `####` |
| 过滤条件（可选） | 去除 `@@@` 行 |

### 标签（Label）使用说明

`@@@标签@@@` 中的内容会被 Dify 识别为该片段的**元数据标签**，可用于：

- 知识库检索时按标签过滤（例如只检索"财务"相关片段）
- 在 Chatflow 中通过 metadata filter 精准召回
- 统计各类别文档数量

标签由「数据标注」工具调用 LLM 自动生成，每个片段最多 5 个标签，每个标签不超过 25 字。

---

## 模型配置

| 用途 | 推荐模型 | 配置位置 |
|------|----------|----------|
| 文档解析（VLM） | `Qwen/Qwen2.5-VL-72B-Instruct` | 右上角「模型设置」 |
| 数据标注（LLM） | `Pro/deepseek-ai/DeepSeek-V3` | 「数据标注」面板内 / API `llm_model` 参数 |

两套 API Key 独立配置，可使用相同或不同的服务商（兼容所有 OpenAI 兼容接口，如硅基流动、OpenAI、Azure 等）。

---

## 目录结构

```
HHPDI文档数据处理智能助手V3/
├── main.py                    # 入口文件（同时启动 GUI 和 API）
├── requirements.txt
├── config/
│   └── settings.py            # 全局配置（API Key、模型等）
├── core/                      # 文档解析核心逻辑
│   ├── pipeline.py            # 主调度器，支持批量并行
│   ├── pdf_loader.py
│   ├── word_loader.py
│   ├── vlm_client.py
│   ├── md_builder.py
│   ├── region_extractor.py
│   └── word_exporter.py
├── gui/                       # 界面框架
│   ├── app.py                 # 主窗口
│   ├── theme.py               # 视觉主题
│   ├── widgets.py             # 共享组件
│   ├── home_panel.py          # 主页
│   └── settings_window.py
├── tools/                     # 各工具面板
│   ├── tool1_parser.py        # 文档解析（批量 + 并行）
│   ├── tool2_converter.py     # MD→Word
│   ├── tool3_annotator.py     # 数据标注（输出 Dify 格式）
│   └── pipeline_panel.py      # 流水线（批量多文件）
└── api/                       # REST API 服务
    ├── server.py              # FastAPI 应用，含全部端点
    ├── job_manager.py         # 异步任务注册与状态追踪
    └── annotator_core.py      # 独立标注核心（无 GUI 依赖）
```
