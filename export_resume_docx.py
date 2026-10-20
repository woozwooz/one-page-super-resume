#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把单文件网页简历（渲染后的 DOM）导出为 Word (.docx)。

用法：
  python3 export_resume_docx.py [源 HTML] [输出 docx]

  默认读取与本脚本同目录的 resume-template.html，输出 resume-template.docx。
  需先安装 playwright 与 python-docx；若本机 Chromium 不在默认位置，
  用环境变量 CHROME_PATH 指定可执行文件。

说明：用 Playwright 渲染页面后按 DOM 结构提取文本（自动跳过被隐藏的楼层、
编辑态 UI、图标等），再用 python-docx 生成 A4 版式的 Word 文档。
"""

import os
import re
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = Path(__file__).resolve().parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "resume-template.html"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else SRC.with_name("resume-template.docx")
# 留空则使用 playwright 自带的 Chromium；需要指定时用 CHROME_PATH 环境变量
CHROME = os.environ.get("CHROME_PATH") or None

INK = "1D1A14"
INK2 = "5C554A"
INK3 = "8C8375"
ACCENT = "A37E46"
FONT = ("Microsoft YaHei", "微软雅黑")

EXTRACT_JS = r"""() => {
  const main = document.getElementById('main');
  if (!main) return { error: 'no #main' };
  const JUNK = 'sider-quick-compose-btn,.js-ui,.mini-del,.add-row,.qr-del,.link-hint,svg,script,style';

  function lines(el) {
    if (!el) return [];
    const c = el.cloneNode(true);
    c.querySelectorAll(JUNK).forEach(n => n.remove());
    c.querySelectorAll('br').forEach(b => b.replaceWith('\n'));
    ['div', 'p', 'li', 'ul', 'h1', 'h2', 'h3', 'h4', 'blockquote'].forEach(t => {
      c.querySelectorAll(t).forEach(b => {
        b.insertAdjacentText('beforebegin', '\n');
        b.insertAdjacentText('afterend', '\n');
      });
    });
    const t = (c.textContent || '').replace(/\u00a0/g, ' ');
    return t.split('\n').map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  }
  const txt = (sel, root) => {
    const el = (root || main).querySelector(sel);
    return el ? lines(el).join(' ') : '';
  };
  const hidden = el => !el || el.classList.contains('floor-hidden') ||
      el.classList.contains('nav-hidden') || getComputedStyle(el).display === 'none';

  const out = { sections: [] };

  // ---------- HERO ----------
  const roles = ['hero.role1', 'hero.role2', 'hero.role3']
      .map(k => txt('[data-edit="' + k + '"]')).filter(Boolean);
  out.hero = {
    name: txt('[data-edit="hero.name"]'),
    nameen: txt('[data-edit="hero.nameen"]'),
    role: roles.join(' / '),
    tel: txt('[data-edit="hero.tel"]'),
    email: txt('[data-edit="hero.email"]'),
    intro: lines(main.querySelector('[data-edit="hero.intro"]')),
  };

  // ---------- 数据条 ----------
  out.metrics = [...main.querySelectorAll('.metric.clone-item')].map(m => ({
    num: txt('.num', m),
    lab: txt('.lab', m),
  }));

  // ---------- 各楼层 ----------
  const secIds = ['strengths', 'ai', 'experience', 'misc', 'project', 'contact'];
  secIds.forEach(id => {
    const sec = main.querySelector('#' + id);
    if (!sec || hidden(sec)) return;
    const item = { id, no: txt('.no', sec), title: txt('h2', sec), en: txt('.en', sec), blocks: [] };

    if (id === 'strengths') {
      sec.querySelectorAll('.s-card').forEach(c => {
        const h3 = c.querySelector('h3');
        const title = lines(h3);
        item.blocks.push({
          kind: 'card',
          title: title[0] || '',
          sub: title.slice(1).join(' '),
          lines: lines(c.querySelector('p')),
        });
      });
    } else if (id === 'ai') {
      sec.querySelectorAll('.ai-node').forEach(n => {
        item.blocks.push({
          kind: 'node',
          year: txt('.year', n),
          tag: txt('.tag', n),
          lines: lines(n.querySelector('p')),
        });
      });
    } else if (id === 'experience') {
      sec.querySelectorAll('.exp-entry').forEach(e => {
        const tags = [...e.querySelectorAll('.exp-tags .clone-item')]
            .map(s => lines(s).join(' ')).filter(Boolean);
        const descEl = e.querySelector(':scope > p');
        const blocks = [...e.querySelectorAll('.exp-block')].map(b => {
          const bt = b.querySelector('.blk-title');
          let title = '', meta = '';
          if (bt) {
            const cc = bt.cloneNode(true);
            const m = cc.querySelector('.meta');
            if (m) { meta = lines(m).join(' '); m.remove(); }
            title = lines(cc).join(' ');
          }
          return {
            title,
            meta,
            items: [...b.querySelectorAll('ul > li.clone-item')].map(li => lines(li)),
          };
        });
        item.blocks.push({
          kind: 'entry',
          period: txt('.exp-period', e),
          org: txt('.exp-org', e),
          role: txt('.exp-role', e),
          tags,
          desc: descEl ? lines(descEl) : [],
          blocks,
        });
      });
    } else if (id === 'misc') {
      sec.querySelectorAll('.misc-card').forEach(c => {
        const h3 = c.querySelector('h3');
        const ic = h3 ? (h3.querySelector('.ic') || {}).textContent || '' : '';
        let title = '';
        if (h3) {
          const cc = h3.cloneNode(true);
          cc.querySelectorAll('.ic').forEach(n => n.remove());
          title = lines(cc).join(' ');
        }
        item.blocks.push({ kind: 'card', ic: (ic || '').trim(), title, lines: lines(c.querySelector('p')) });
      });
    } else {
      item.blocks.push({ kind: 'card', title: '', sub: '', lines: lines(sec) });
    }
    out.sections.push(item);
  });

  return out;
}
"""


def extract():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            **({"executable_path": CHROME} if CHROME else {}))
        page = browser.new_page()
        page.goto(SRC.as_uri(), wait_until="domcontentloaded", timeout=60000)
        data = page.evaluate(EXTRACT_JS)
        browser.close()
    if not data or data.get("error"):
        raise RuntimeError(f"提取失败: {data}")
    return data


# ---------------- docx 工具 ----------------

def set_font(run, size=10, bold=False, color=None, italic=False):
    run.font.name = FONT[0]
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), FONT[1])
    rFonts.set(qn("w:ascii"), FONT[0])
    rFonts.set(qn("w:hAnsi"), FONT[0])
    return run


def para(doc, space_after=2, space_before=0, indent=None, line=1.1):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line
    if indent is not None:
        pf.left_indent = Cm(indent)
    return p


def runs(p, parts):
    """parts: [(text, size, bold, color)]"""
    for text, size, bold, color in parts:
        if not text:
            continue
        set_font(p.add_run(text), size=size, bold=bold, color=color)
    return p


def text_para(doc, text, size=10, color=INK2, bold=False, indent=None,
              space_after=2, line=1.1):
    p = para(doc, space_after=space_after, indent=indent, line=line)
    runs(p, [(text, size, bold, color)])
    return p


def hr(doc, color=ACCENT, sz=8, before=0, after=4):
    p = para(doc, space_after=after, space_before=before, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), str(sz))
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color)
    pbdr.append(bot)
    pPr.append(pbdr)
    return p


def no_border_table(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        borders.append(el)
    tblPr.append(borders)


def clean_item(text):
    t = re.sub(r"^[\-–•]\s*", "", text.strip())
    return t


def clean_sub(text):
    t = re.sub(r"^[>＞]\s*", "", text.strip())
    return t


def bullet(doc, text, level=0, size=9.5, color=INK2, bold=False):
    indent = 0.45 + 0.45 * level
    p = para(doc, space_after=1, indent=indent, line=1.08)
    pf = p.paragraph_format
    pf.first_line_indent = Cm(-0.32)
    set_font(p.add_run("· " if level == 0 else "– "), size=size, color=color)
    set_font(p.add_run(text), size=size, color=color, bold=bold)
    return p


# ---------------- 组装 ----------------

def build(data):
    doc = Document()
    sec0 = doc.sections[0]
    sec0.top_margin = Cm(1.7)
    sec0.bottom_margin = Cm(1.7)
    sec0.left_margin = Cm(2.0)
    sec0.right_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = FONT[0]
    normal.font.size = Pt(10)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT[1])

    hero = data["hero"]

    # 姓名
    p = para(doc, space_after=1, line=1.0)
    runs(p, [(hero["name"], 20, True, INK), ("   ", 20, False, INK),
             (hero["nameen"], 13, False, INK3)])
    # 角色行
    if hero["role"]:
        p = para(doc, space_after=2, line=1.0)
        runs(p, [(hero["role"], 10.5, True, ACCENT)])
    # 联系方式
    contact = "    ｜    ".join([x for x in [hero["tel"], hero["email"]] if x])
    if contact:
        p = para(doc, space_after=3, line=1.0)
        runs(p, [("电话 ", 9, False, INK3), (hero["tel"], 9.5, False, INK2),
                 ("    ｜    ", 9, False, INK3), ("邮箱 ", 9, False, INK3),
                 (hero["email"], 9.5, False, INK2)])
    hr(doc, sz=10, after=6)

    # 简介
    for line in hero["intro"]:
        text_para(doc, line, size=10, color=INK2, space_after=1)

    # 数据条
    metrics = [m for m in data["metrics"] if m["num"] or m["lab"]]
    if metrics:
        hr(doc, sz=4, before=4, after=4)
        n = len(metrics)
        table = doc.add_table(rows=1, cols=n)
        table.alignment = 1
        no_border_table(table)
        width = Cm((21.0 - 4.0) / n)
        for i, m in enumerate(metrics):
            cell = table.cell(0, i)
            cell.width = width
            cp = cell.paragraphs[0]
            cp.alignment = 1
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.line_spacing = 1.0
            set_font(cp.add_run(m["num"]), size=13, bold=True, color=ACCENT)
            lp = cell.add_paragraph()
            lp.alignment = 1
            lp.paragraph_format.space_after = Pt(0)
            lp.paragraph_format.space_before = Pt(1)
            lp.paragraph_format.line_spacing = 1.0
            set_font(lp.add_run(m["lab"]), size=8, color=INK3)
        hr(doc, sz=4, before=4, after=2)

    # 楼层
    for s in data["sections"]:
        p = para(doc, space_before=8, space_after=2, line=1.0)
        runs(p, [(s["no"] + "  ", 10, True, ACCENT),
                 (s["title"], 13, True, INK),
                 ("    " + s["en"] if s["en"] else "", 9, False, INK3)])
        hr(doc, sz=6, after=5)

        for b in s["blocks"]:
            if b["kind"] == "card":
                if b.get("title"):
                    p = para(doc, space_before=2, space_after=1, line=1.05)
                    if b.get("ic"):
                        runs(p, [(b["ic"] + "  ", 9, True, ACCENT)])
                    runs(p, [(b["title"], 10.5, True, INK)])
                    if b.get("sub"):
                        runs(p, [("   " + b["sub"], 9, False, INK3)])
                for line in b.get("lines", []):
                    if s["id"] == "misc":
                        text_para(doc, clean_item(line), size=10, indent=0.2)
                    else:
                        bullet(doc, clean_item(line))
            elif b["kind"] == "node":
                p = para(doc, space_before=2, space_after=1, line=1.05)
                runs(p, [(b["year"], 10.5, True, ACCENT),
                         ("    " + b["tag"] if b["tag"] else "", 9, True, INK3)])
                for line in b.get("lines", []):
                    bullet(doc, clean_item(line))
            elif b["kind"] == "entry":
                p = para(doc, space_before=6, space_after=1, line=1.05)
                runs(p, [(b["period"] + "    ", 9.5, False, INK3),
                         (b["org"], 11.5, True, INK),
                         ("    " + b["role"] if b["role"] else "", 10, False, ACCENT)])
                if b["tags"]:
                    p = para(doc, space_after=1, line=1.0)
                    runs(p, [("   ".join(b["tags"]), 9, False, INK3)])
                for line in b.get("desc", []):
                    text_para(doc, line, size=10, color=INK2, space_after=1, indent=0.2)
                for blk in b.get("blocks", []):
                    if blk["title"]:
                        p = para(doc, space_before=3, space_after=1, line=1.05)
                        runs(p, [(blk["title"], 10, True, INK),
                                 ("    " + blk.get("meta", "") if blk.get("meta") else "",
                                  9, False, INK3)])
                    for item in blk["items"]:
                        if not item:
                            continue
                        bullet(doc, clean_item(item[0]))
                        for sub in item[1:]:
                            if sub:
                                bullet(doc, clean_sub(sub), level=1)
            else:
                for line in b.get("lines", []):
                    text_para(doc, line)

    doc.core_properties.title = f"{hero['name']} {hero['nameen']} 的简历"
    doc.save(OUT)
    return doc


def main():
    data = extract()
    build(data)
    print(f"已生成: {OUT}  ({OUT.stat().st_size/1024:.1f} KB)")
    print("楼层:", [s["title"] for s in data["sections"]])
    print("数据条:", len(data["metrics"]), "| 经历段数:",
          sum(len(s["blocks"]) for s in data["sections"] if s["id"] == "experience"))


if __name__ == "__main__":
    main()
