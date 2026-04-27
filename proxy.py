#!/usr/bin/env python3
"""
WPS AI助手 - 本地代理服务器
功能：
  1. 静态文件服务（ribbon.xml / main.js / taskpane.*）
  2. /api/chat  → 转发到内网 AI 平台（HTTPS 自签名证书绕过）
  3. /api/kb    → 转发到智弈智枢知识库
端口：8765
"""

import os
import json
import sys
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import httpx
from dotenv import load_dotenv

# ── 加载环境变量（优先读同目录下的 .env）──────────────────────
_here = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(_here, ".env")
if not os.path.exists(ENV_PATH):
    # 兼容开发环境
    ENV_PATH = r"D:\claude\training-manager\.env"
load_dotenv(ENV_PATH, override=True)

AIRCHINA_API_KEY  = os.getenv("AIRCHINA_API_KEY",  "")
AIRCHINA_BASE_URL = os.getenv("AIRCHINA_BASE_URL", "")
ZHISHU_API_KEY    = os.getenv("ZHISHU_API_KEY",    "")
ZHISHU_BASE_URL   = os.getenv("ZHISHU_BASE_URL",   "")
MODEL_CHAT        = os.getenv("MODEL_CHAT",         "glm-5-outside")

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 8765

# ── CORS 头 ──────────────────────────────────────────────────────
CORS = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
}


class Handler(SimpleHTTPRequestHandler):
    """静态文件 + API 代理 合一"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PROJECT_DIR, **kwargs)

    # ── 静默日志（避免刷屏）─────────────────────────────────────
    def log_message(self, fmt, *args):
        path = args[0] if args else ""
        # 只打印 API 调用，忽略静态资源请求
        if "/api/" in str(path):
            print(f"[API] {fmt % args}")

    # ── 添加 CORS 头 ─────────────────────────────────────────────
    def end_headers(self):
        for k, v in CORS.items():
            self.send_header(k, v)
        super().end_headers()

    # ── OPTIONS 预检 ─────────────────────────────────────────────
    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    # ── GET：静态文件（SimpleHTTPRequestHandler 已处理）─────────
    # 仅对 /api/status 单独处理；静态文件需去掉查询参数（?v=N 缓存破坏）
    def do_GET(self):
        if self.path == "/api/status":
            self._json(200, {"status": "ok", "model": MODEL_CHAT})
        else:
            # 剥离 ?v=N 等查询参数，防止 SimpleHTTPRequestHandler 找不到文件
            if "?" in self.path and not self.path.startswith("/api/"):
                self.path = self.path.split("?")[0]
            super().do_GET()

    # ── POST：API 路由 ────────────────────────────────────────────
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            body = json.loads(self.rfile.read(length)) if length else {}
        except Exception:
            body = {}

        if self.path == "/api/chat":
            self._handle_chat(body)
        elif self.path == "/api/kb":
            self._handle_kb(body)
        else:
            self._json(404, {"error": "not found"})

    # ── 普通对话：转发到 AIRCHINA OpenAI 兼容接口 ───────────────
    def _handle_chat(self, body):
        try:
            messages = body.get("messages", [])
            url      = f"{AIRCHINA_BASE_URL}/chat/completions"
            headers  = {
                "Authorization": f"Bearer {AIRCHINA_API_KEY}",
                "Content-Type":  "application/json",
            }
            payload = {
                "model":    MODEL_CHAT,
                "messages": messages,
            }
            resp = httpx.post(
                url, json=payload, headers=headers,
                verify=False, timeout=120,
            )
            resp.raise_for_status()
            self._json(200, resp.json())
        except Exception as e:
            self._json(500, {"error": str(e)})

    # ── 知识库问答：转发到智弈智枢 ───────────────────────────────
    def _handle_kb(self, body):
        try:
            question        = body.get("query", "")
            conversation_id = body.get("conversation_id", "")
            url     = f"{ZHISHU_BASE_URL}/chat-messages"
            headers = {
                "Authorization": f"Bearer {ZHISHU_API_KEY}",
                "Content-Type":  "application/json",
            }
            payload = {
                "query":           question,
                "inputs":          {},
                "response_mode":   "blocking",
                "user":            "wps-plugin",
                "conversation_id": conversation_id,
            }
            resp = httpx.post(
                url, json=payload, headers=headers,
                verify=False, timeout=120,
            )
            resp.raise_for_status()
            self._json(200, resp.json())
        except Exception as e:
            self._json(500, {"error": str(e)})

    # ── 工具：返回 JSON ──────────────────────────────────────────
    def _json(self, status: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    # 屏蔽 httpx 的 InsecureRequestWarning
    import warnings
    import urllib3
    warnings.filterwarnings("ignore")
    try:
        urllib3.disable_warnings()
    except Exception:
        pass

    server = HTTPServer(("127.0.0.1", PORT), Handler)

    print("=" * 55)
    print("  WPS AI助手 代理服务器")
    print("=" * 55)
    print(f"  地址     : http://127.0.0.1:{PORT}")
    print(f"  AI模型   : {MODEL_CHAT}")
    print(f"  知识库   : {ZHISHU_BASE_URL}")
    print(f"  项目目录 : {PROJECT_DIR}")
    print("=" * 55)
    print("  按 Ctrl+C 停止服务")
    print()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
        server.shutdown()


if __name__ == "__main__":
    main()
