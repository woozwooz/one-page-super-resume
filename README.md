# 单文件网页简历模板

一个**零依赖、单 HTML 文件**的中文简历模板：打开即渲染，页面上直接编辑文字，退出编辑态自动回写文件；另配一个脚本把渲染结果导出成 A4 版式的 Word。

- `resume-template.html` —— 模板本体（内含编辑系统、打印样式、PDF 精简版式）。示例内容为占位信息（陈亦舟 / Ethan），改成自己的即可。
- `resume_sync.py` —— 本地静态服务 + 保存回写（编辑用）。
- `export_resume_docx.py` —— 把渲染后的页面导出为 `.docx`。

## 快速开始

### 1. 编辑

编辑功能本身不需要任何依赖。把改动**写回文件**有两条路，页面会自动选：

1. **直接连原文件（推荐，零依赖）**——用 Chrome / Edge 打开，第一次点「编辑」时浏览器会让你选一次本文件，之后每次点「完成 / 退出编辑态」就直接写回磁盘，不需要任何服务。授权记在浏览器里，下次打开仍然有效。
2. **本地同步服务**——见下。服务开着时**不会**弹文件选择框，页面静默走服务回写。

两条都不通（比如 Safari / Firefox 且没起服务）时，自动降级为**下载一份副本**，内容不会丢。

> 浏览器出于安全限制，无法在打开 HTML 时自动启动本机程序，所以"完全无感"只能做到「授权一次 → 之后静默」。

#### 本地同步服务（可选）

```bash
python3 resume_sync.py                 # 默认 8765 端口，托管本目录
python3 resume_sync.py --port 9000     # 换端口
python3 resume_sync.py --dir /path     # 托管别的目录
```

然后浏览器打开终端里列出的地址（默认 <http://127.0.0.1:8765/resume-template.html>），直接点文字改。
点击「完成 / 退出编辑态」时，页面会把当前完整 HTML POST 回 `/__save`，服务端原子写回原文件，并在写前留一份备份到 `.resume-backups/`（同名最多保留 10 份）。

服务已放开 CORS，所以**用 `file://` 双击打开页面、服务在后台跑着**也能自动回写，不一定要从服务地址打开。

不启动服务、浏览器也不支持直连文件（Safari / Firefox）时，仍能正常编辑，内容存在浏览器 localStorage 里，退出编辑态会导出一份副本文件。

### 2. 导出 Word

编辑本身不需要任何依赖，只有这个导出脚本需要两个 Python 包：`playwright>=1.44`、`python-docx>=1.1.0`。

```bash
pip install "playwright>=1.44" "python-docx>=1.1.0"
playwright install chromium            # 或设 CHROME_PATH 指向已有 Chromium
python3 export_resume_docx.py                        # 默认 resume-template.html → resume-template.docx
python3 export_resume_docx.py 源文件.html 输出.docx   # 指定文件名
```

导出用 Playwright 渲染页面后按 DOM 提取文本（自动跳过隐藏楼层、编辑态 UI、图标），生成 A4、微软雅黑、带金色分隔线的 docx。

### 3. 导出 PDF

直接用浏览器打印（`Ctrl/Cmd + P`）即可，模板带专门的打印样式与「精简版式」开关。

## 安全提醒

`resume_sync.py` **只监听 `127.0.0.1`**，并且有写文件接口。请不要把它部署到公网或改成 `0.0.0.0` 暴露出去。

它的写入限制（本地防护）：只允许写托管目录内**已存在**的 `.html`、文件名不得含路径分隔符或 `..`、单次请求体上限 20MB、写盘为临时文件 + `os.replace` 原子替换。

## 已知注意事项

- **字体**：docx 导出写死「微软雅黑 / Microsoft YaHei」。Windows 正常；macOS / Linux 上生成的文档在缺少该字体的机器打开会回退字体，需要改的话调整 `export_resume_docx.py` 顶部的 `FONT` 常量。
- **在线字体**：HTML 从 Google Fonts 加载 Noto Serif SC / Inter / JetBrains Mono。离线或访问不畅时排版会降级回系统字体（不影响功能）；介意访客 IP 泄露给 Google 的话，可自行改为本地字体或删除 `<link>`。
- **备份目录**：`.resume-backups/` 由 `resume_sync.py` 自动创建，建议一并加入你的本地忽略规则（本仓库已用白名单式 `.gitignore` 排除）。

## 许可

代码 MIT，见 [LICENSE](./LICENSE)。模板中的示例文案为占位内容，可自由改写使用。
