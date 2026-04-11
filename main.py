#!/usr/bin/env python3
"""
DocFlow Pro — 入口文件
集成：文档解析（DocFlow）+ MD→Word（MinerU Converter）+ MD数据标注
附加：REST API 服务（可选，需安装 fastapi + uvicorn）
"""
import sys
import os

# 确保项目根目录在模块搜索路径
_root = os.path.dirname(os.path.abspath(__file__))
if _root not in sys.path:
    sys.path.insert(0, _root)


def _check_deps():
    """检查必要依赖"""
    missing = []
    optional_missing = []

    # 必须有 tkinter
    try:
        import tkinter
    except ImportError:
        print("错误：找不到 tkinter，请使用 python.org 官方安装包。")
        input("按回车退出…")
        sys.exit(1)

    # 核心依赖（Tool 2 需要）
    try:
        from docx import Document
    except ImportError:
        optional_missing.append("python-docx")

    try:
        from bs4 import BeautifulSoup
    except ImportError:
        optional_missing.append("beautifulsoup4")

    # Tool 1 依赖
    try:
        import fitz
    except ImportError:
        optional_missing.append("pymupdf")

    try:
        import openai
    except ImportError:
        optional_missing.append("openai")

    # Tool 3 依赖
    try:
        import requests
    except ImportError:
        optional_missing.append("requests")

    if optional_missing:
        try:
            import tkinter as tk
            from tkinter import messagebox
            _r = tk.Tk(); _r.withdraw()
            messagebox.showwarning(
                "部分依赖缺失",
                f"以下依赖未安装，相关功能可能不可用：\n\n"
                f"  pip install {' '.join(optional_missing)}\n\n"
                "基础功能仍可使用，安装后重启以获得完整功能。"
            )
            _r.destroy()
        except Exception:
            print(f"警告：缺少依赖: {', '.join(optional_missing)}")
            print(f"运行: pip install {' '.join(optional_missing)}")


def _start_api_server():
    """
    在独立后台线程中启动 REST API 服务器。
    - 使用 uvicorn 的编程接口，避免与 Tkinter 主线程的事件循环冲突
    - 若 fastapi / uvicorn 未安装则静默跳过，不影响桌面端
    """
    try:
        import asyncio
        import uvicorn
        from api.server import app as api_app

        host = os.environ.get("HHPDI_API_HOST", "127.0.0.1")
        port = int(os.environ.get("HHPDI_API_PORT", "8765"))

        # 在本线程内建立独立事件循环（与 Tkinter 主线程隔离）
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        config = uvicorn.Config(
            api_app,
            host=host,
            port=port,
            loop="asyncio",
            log_level="warning",
            use_colors=False,
        )
        server = uvicorn.Server(config)
        print(f"[HHPDI API] 服务已启动 → http://{host}:{port}/docs")
        loop.run_until_complete(server.serve())
    except ImportError:
        pass  # fastapi / uvicorn 未安装，跳过
    except OSError as e:
        print(f"[HHPDI API] 端口被占用或无法绑定: {e}")
    except Exception as e:
        print(f"[HHPDI API] 启动失败: {e}")


def main():
    _check_deps()

    # 后台启动 API 服务（daemon=True：GUI 关闭时自动终止）
    import threading
    api_thread = threading.Thread(target=_start_api_server, daemon=True)
    api_thread.start()

    try:
        from gui.app import DocFlowProApp
        app = DocFlowProApp()
        app.mainloop()
    except Exception:
        import traceback
        err = traceback.format_exc()
        try:
            import tkinter as tk
            from tkinter import messagebox
            _r = tk.Tk(); _r.withdraw()
            messagebox.showerror("启动错误", err[:600])
        except Exception:
            print(f"启动错误：\n{err}")
            input("按回车退出…")


if __name__ == "__main__":
    main()
