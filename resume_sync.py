#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
resume_sync.py —— 简历「保存即回写文件」的本地同步服务

做两件事：
  1) 以 http://127.0.0.1:<port> 静态托管简历目录（同源，页面 fetch 无需 CORS）
  2) POST /__save 接收页面烘焙好的完整 HTML，原子写回同名文件（写前自动备份）
  3) POST /__save_as 把烘焙好的 HTML「另存副本」到本目录（新建文件，
     文件名由页面给出、尾缀带日期时间；同名已存在自动 -2 / -3 避让，绝不覆盖）

用法：
    python3 resume_sync.py                # 默认 8765 端口，托管本文件所在目录
    python3 resume_sync.py --port 9000
    python3 resume_sync.py --dir /path/to/resume

然后浏览器打开 http://127.0.0.1:8765/resume-template.html 编辑，
「完成 / 退出编辑态」时页面会 POST 回写，终端可见日志。

安全约束（全部为本地防护，服务只绑定 127.0.0.1）：
  - 只允许写目录内已存在的 .html 文件，文件名不得含路径分隔符或 ..
  - 单次请求体上限 20MB
  - 写盘 = 临时文件 + os.replace 原子替换，写前留 .bak-时间戳，最多保留 10 份
"""

import argparse
import glob
import json
import os
import re
import sys
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

DEFAULT_PORT = 8765
MAX_BODY = 20 * 1024 * 1024      # 20MB
BACKUP_KEEP = 10
SAFE_NAME = re.compile(r"^[^/\\\x00]+$")

ROOT = ""      # 启动时赋值（--dir 或本文件所在目录）


def log(msg):
    stamp = datetime.now().strftime("%H:%M:%S")
    print("[%s] %s" % (stamp, msg), flush=True)


def resolve_target(name):
    """把请求里的文件名解析成目录内的真实路径；非法返回 (None, 原因)"""
    if not name or not SAFE_NAME.match(name):
        return None, "非法文件名"
    if ".." in name:
        return None, "文件名不得包含 .."
    if not name.lower().endswith(".html"):
        return None, "只允许写 .html 文件"
    target = os.path.realpath(os.path.join(ROOT, name))
    if os.path.dirname(target) != os.path.realpath(ROOT):
        return None, "越出简历目录"
    if not os.path.isfile(target):
        return None, "目标文件不存在：%s" % name
    return target, None


def resolve_new_target(name):
    """「另存副本」：目录内新建一个 .html；同名已存在则自动 -2 / -3 避让，绝不覆盖已有文件。
    非法返回 (None, 原因)"""
    if not name or not SAFE_NAME.match(name):
        return None, "非法文件名"
    if ".." in name:
        return None, "文件名不得包含 .."
    if name.startswith("."):
        return None, "不允许隐藏文件"
    if not name.lower().endswith(".html"):
        return None, "只允许写 .html 文件"
    target = os.path.realpath(os.path.join(ROOT, name))
    if os.path.dirname(target) != os.path.realpath(ROOT):
        return None, "越出简历目录"
    if os.path.exists(target):                       # 同名避让，避免覆盖母本或既有副本
        stem, ext = os.path.splitext(target)
        i = 2
        while os.path.exists("%s-%d%s" % (stem, i, ext)):
            i += 1
        target = "%s-%d%s" % (stem, i, ext)
    return target, None


def backup_dir():
    """备份放在隐藏子目录，避免 .bak-* 堆在简历文件夹里碍眼"""
    d = os.path.join(ROOT, ".resume-backups")
    os.makedirs(d, exist_ok=True)
    return d


def prune_backups(name):
    """同名文件只保留最近 BACKUP_KEEP 份备份"""
    pattern = glob.glob(os.path.join(backup_dir(), name + ".bak-*"))
    pattern.sort(reverse=True)
    for old in pattern[BACKUP_KEEP:]:
        try:
            os.remove(old)
        except OSError:
            pass


class Handler(SimpleHTTPRequestHandler):
    server_version = "resume-sync/1.0"

    # 允许 file:// 直接打开的页面也调用本服务（跨域 + 预检），
    # 这样即使不通过 http://127.0.0.1 打开，「退出编辑态自动回写」同样生效。
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Content-Length", "0")
        self.end_headers()

    # ---------- 工具 ----------
    def _json(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    # 文本类型补 charset，避免中文文件名/内容乱码
    def guess_type(self, path):
        t = super().guess_type(path)
        if t.startswith("text/") or t in ("application/javascript", "application/json"):
            return t + "; charset=utf-8"
        return t

    def log_message(self, fmt, *args):      # 静音默认访问日志，只留我们自己的输出
        pass

    # ---------- 探测 ----------
    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path == "/__ping":
            self._json(200, {"ok": True, "dir": ROOT, "pid": os.getpid()})
            return
        if path == "/__files":
            names = sorted(
                n for n in os.listdir(ROOT)
                if n.lower().endswith(".html") and not n.startswith(".")
            )
            self._json(200, {"ok": True, "files": names})
            return
        return super().do_GET()

    # ---------- 回写 / 另存 ----------
    def do_POST(self):
        path = self.path.split("?", 1)[0]
        if path not in ("/__save", "/__save_as"):
            self._json(404, {"ok": False, "error": "not found"})
            return

        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0 or length > MAX_BODY:
            self._json(400, {"ok": False, "error": "请求体为空或过大"})
            return

        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as exc:                      # noqa: BLE001
            self._json(400, {"ok": False, "error": "JSON 解析失败：%s" % exc})
            return

        name = payload.get("file") or ""
        html = payload.get("html") or ""
        if not html:
            self._json(400, {"ok": False, "error": "html 为空，拒绝写入（避免写坏文件）"})
            return

        if path == "/__save_as":                      # 另存副本：新建文件，不覆盖任何已有文件
            target, err = resolve_new_target(name)
            if err:
                log("✗ 拒绝另存 %s：%s" % (name, err))
                self._json(400, {"ok": False, "error": err})
                return
            self._write_new(target, html)
            return

        target, err = resolve_target(name)
        if err:
            log("✗ 拒绝写入 %s：%s" % (name, err))
            self._json(400, {"ok": False, "error": err})
            return

        try:
            with open(target, "r", encoding="utf-8") as fh:
                current = fh.read()
        except OSError as exc:
            self._json(500, {"ok": False, "error": "读取原文件失败：%s" % exc})
            return

        if current == html:
            log("· 无变化，跳过：%s" % name)
            self._json(200, {"ok": True, "skipped": True, "file": name})
            return

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = os.path.join(backup_dir(), "%s.bak-%s" % (name, stamp))
        suffix = 1
        while os.path.exists(backup):                 # 同秒冲突时加序号
            backup = "%s.bak-%s-%d" % (target, stamp, suffix)
            suffix += 1

        tmp = target + ".__tmp_write"
        try:
            with open(backup, "w", encoding="utf-8") as fh:
                fh.write(current)
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write(html)
            os.replace(tmp, target)                   # 原子替换
        except OSError as exc:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            log("✗ 写入失败 %s：%s" % (name, exc))
            self._json(500, {"ok": False, "error": "写入失败：%s" % exc})
            return

        prune_backups(name)
        log("✓ 已写入 %s（%d 字节，备份 .resume-backups/%s）"
            % (name, len(html.encode("utf-8")), os.path.basename(backup)))
        self._json(200, {
            "ok": True,
            "file": name,
            "bytes": len(html.encode("utf-8")),
            "backup": os.path.basename(backup),
        })

    # ---------- 另存副本（新建文件，原子写，不做备份） ----------
    def _write_new(self, target, html):
        size = len(html.encode("utf-8"))
        tmp = target + ".__tmp_write"
        try:
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write(html)
            os.replace(tmp, target)                   # 原子替换
        except OSError as exc:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            log("✗ 另存失败 %s：%s" % (os.path.basename(target), exc))
            self._json(500, {"ok": False, "error": "写入失败：%s" % exc})
            return
        log("✓ 已另存副本 %s（%d 字节）" % (os.path.basename(target), size))
        self._json(200, {
            "ok": True,
            "file": os.path.basename(target),
            "path": target,
            "bytes": size,
        })


def main():
    global ROOT
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description="简历本地同步服务")
    ap.add_argument("--dir", default=here, help="托管的简历目录（默认本文件所在目录）")
    ap.add_argument("--port", type=int, default=DEFAULT_PORT, help="监听端口（默认 8765）")
    args = ap.parse_args()

    ROOT = os.path.realpath(args.dir)
    if not os.path.isdir(ROOT):
        print("目录不存在：%s" % ROOT, file=sys.stderr)
        return 2

    Handler.directory = ROOT
    srv = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)

    files = sorted(n for n in os.listdir(ROOT) if n.lower().endswith(".html") and not n.startswith("."))
    log("简历同步服务已启动 → http://127.0.0.1:%d/" % args.port)
    log("目录：%s" % ROOT)
    for n in files:
        log("  · http://127.0.0.1:%d/%s" % (args.port, n.replace(" ", "%20")))
    log("退出编辑态时页面会自动 POST /__save 回写；Ctrl+C 停止")

    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        log("已停止")
    finally:
        srv.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
