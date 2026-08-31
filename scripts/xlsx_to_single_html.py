#!/usr/bin/env python3
"""把千川拆解 Excel 转成包含内嵌图片的单文件 HTML。"""

import argparse
import base64
import html
import json
import mimetypes
import re
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

HTML_TEMPLATE_VERSION = "1.0.0"


def rel_path(source, target):
    source_dir = PurePosixPath(source).parent
    parts = []
    for part in (source_dir / target).parts:
        if part == "..":
            if parts:
                parts.pop()
        elif part not in (".", "/"):
            parts.append(part)
    return "/".join(parts)


def relationship_map(book, rels_path, source_path):
    if rels_path not in book.namelist():
        return {}
    root = ET.fromstring(book.read(rels_path))
    return {
        rel.attrib["Id"]: rel_path(source_path, rel.attrib["Target"])
        for rel in root.findall("rel:Relationship", NS)
    }


def shared_strings(book):
    path = "xl/sharedStrings.xml"
    if path not in book.namelist():
        return []
    root = ET.fromstring(book.read(path))
    return ["".join(node.text or "" for node in item.findall(".//m:t", NS)) for item in root.findall("m:si", NS)]


def column_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    result = 0
    for char in letters:
        result = result * 26 + ord(char) - 64
    return result - 1


def cell_value(cell, strings):
    cell_type = cell.attrib.get("t")
    value = cell.find("m:v", NS)
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//m:t", NS))
    if value is None:
        return ""
    raw = value.text or ""
    if cell_type == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "b":
        return "是" if raw == "1" else "否"
    if cell_type in ("str", "e"):
        return raw
    try:
        number = float(raw)
        return str(int(number)) if number.is_integer() else f"{number:g}"
    except ValueError:
        return raw


def drawing_images(book, sheet_path, sheet_root):
    result = {}
    drawing = sheet_root.find("m:drawing", NS)
    if drawing is None:
        return result
    sheet_rels_path = str(PurePosixPath(sheet_path).parent / "_rels" / (PurePosixPath(sheet_path).name + ".rels"))
    sheet_rels = relationship_map(book, sheet_rels_path, sheet_path)
    drawing_path = sheet_rels.get(drawing.attrib.get(f"{{{NS['r']}}}id"))
    if not drawing_path or drawing_path not in book.namelist():
        return result
    drawing_root = ET.fromstring(book.read(drawing_path))
    drawing_rels_path = str(PurePosixPath(drawing_path).parent / "_rels" / (PurePosixPath(drawing_path).name + ".rels"))
    drawing_rels = relationship_map(book, drawing_rels_path, drawing_path)
    for anchor in list(drawing_root.findall("xdr:oneCellAnchor", NS)) + list(drawing_root.findall("xdr:twoCellAnchor", NS)):
        start = anchor.find("xdr:from", NS)
        image = anchor.find(".//a:blip", NS)
        if start is None or image is None:
            continue
        row = int(start.findtext("xdr:row", default="0", namespaces=NS))
        col = int(start.findtext("xdr:col", default="0", namespaces=NS))
        media_path = drawing_rels.get(image.attrib.get(f"{{{NS['r']}}}embed"))
        if not media_path or media_path not in book.namelist():
            continue
        mime = mimetypes.guess_type(media_path)[0] or "image/jpeg"
        encoded = base64.b64encode(book.read(media_path)).decode("ascii")
        result[(row, col)] = f"data:{mime};base64,{encoded}"
    return result


def parse_workbook(path):
    with zipfile.ZipFile(path) as book:
        strings = shared_strings(book)
        workbook_root = ET.fromstring(book.read("xl/workbook.xml"))
        workbook_rels = relationship_map(book, "xl/_rels/workbook.xml.rels", "xl/workbook.xml")
        sheets = []
        for sheet in workbook_root.findall("m:sheets/m:sheet", NS):
            sheet_path = workbook_rels[sheet.attrib[f"{{{NS['r']}}}id"]]
            root = ET.fromstring(book.read(sheet_path))
            image_map = drawing_images(book, sheet_path, root)
            parsed_rows = []
            max_col = 0
            for row in root.findall("m:sheetData/m:row", NS):
                row_index = int(row.attrib.get("r", len(parsed_rows) + 1)) - 1
                while len(parsed_rows) <= row_index:
                    parsed_rows.append([])
                values = parsed_rows[row_index]
                for cell in row.findall("m:c", NS):
                    col = column_index(cell.attrib["r"])
                    while len(values) <= col:
                        values.append("")
                    values[col] = cell_value(cell, strings)
                    max_col = max(max_col, col + 1)
            if image_map:
                max_col = max(max_col, max(col for _, col in image_map) + 1)
                max_row = max(row for row, _ in image_map)
                while len(parsed_rows) <= max_row:
                    parsed_rows.append([])
            for row in parsed_rows:
                row.extend([""] * (max_col - len(row)))
            sheets.append({"name": sheet.attrib["name"], "rows": parsed_rows, "images": image_map})
        return sheets


def safe_text(value):
    return html.escape(str(value or "")).replace("\n", "<br>")


def sheet_stats(sheets):
    overview = next((s for s in sheets if s["name"] == "素材总览"), None)
    evidence = next((s for s in sheets if s["name"] == "画面证据"), None)
    materials = max(0, len(overview["rows"]) - 4) if overview else 0
    images = len(evidence["images"]) if evidence else sum(len(s["images"]) for s in sheets)
    return materials, images


def sheet_records(sheet):
    rows = sheet["rows"]
    if len(rows) < 4:
        return [], "", ""
    headers = rows[3]
    records = []
    for row_index, row in enumerate(rows[4:], 4):
        record = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers)) if headers[i]}
        record["__row_index"] = row_index
        records.append(record)
    title = rows[0][0] if rows and rows[0] else sheet["name"]
    note = rows[1][0] if len(rows) > 1 and rows[1] else ""
    return records, title, note


def detail(label, value, tone=""):
    if value in (None, ""):
        return ""
    return f'<div class="detail {tone}"><dt>{safe_text(label)}</dt><dd>{safe_text(value)}</dd></div>'


def overview_section(sheet):
    records, title, note = sheet_records(sheet)
    cards = []
    for item in records:
        chain = "".join(f"<span>{safe_text(node.strip())}</span>" for node in str(item.get("一句话成交链", "")).split("→") if node.strip())
        meta = " · ".join(filter(None, [item.get("原始文件名"), item.get("时长（秒）") and f'{item.get("时长（秒）")}秒', item.get("画幅"), item.get("产品/品类")]))
        cards.append(f'''<article class="overview-card">
          <div class="card-top"><div><span class="badge">素材 {safe_text(item.get('素材编号'))}</span><h3>{safe_text(item.get('成交模型'))}</h3><p class="meta">{safe_text(meta)}</p></div><div class="priority">{safe_text(item.get('优先级'))}<small>优先级</small></div></div>
          <div class="hook"><b>核心钩子</b><p>{safe_text(item.get('核心钩子'))}</p></div>
          <div class="chain" aria-label="一句话成交链：{safe_text(item.get('一句话成交链'))}"><b>一句话成交链</b><div class="chain-flow">{chain}</div></div>
          <dl class="detail-grid">{detail('原始文件名',item.get('原始文件名'))}{detail('时长（秒）',item.get('时长（秒）'))}{detail('画幅',item.get('画幅'))}{detail('产品/品类',item.get('产品/品类'))}{detail('跑量逻辑',item.get('跑量逻辑'))}{detail('最大优点',item.get('最大优点'),'good')}{detail('主要不足',item.get('主要不足'),'warn')}{detail('合规风险',item.get('合规风险'),'risk')}{detail('模型状态',item.get('模型状态'))}{detail('模型置信度',item.get('模型置信度'))}{detail('建议素材池角色',item.get('建议素材池角色'))}{detail('画面证据编号',item.get('画面证据编号'))}{detail('来源目录',item.get('来源目录'))}</dl>
        </article>''')
    return f'<section id="overview"><div class="section-head"><p>01 · 核心结论</p><h2>{safe_text(title)}</h2><span>{safe_text(note)}</span></div>{"".join(cards)}</section>'


def segments_section(sheet):
    records, _, note = sheet_records(sheet)
    items = []
    for item in records:
        items.append(f'''<article class="timeline-item"><div class="time"><b>{safe_text(item.get('起始时间'))}–{safe_text(item.get('结束时间'))}s</b><span>{safe_text(item.get('内容功能'))}</span></div><div class="timeline-card"><div class="timeline-title"><h3>内容功能：{safe_text(item.get('内容功能'))}</h3><span>画面证据编号：{safe_text(item.get('画面证据编号'))}</span></div><p class="segment-meta">原始文件名：{safe_text(item.get('原始文件名'))} · 起始时间：{safe_text(item.get('起始时间'))} · 结束时间：{safe_text(item.get('结束时间'))}</p><p class="scene"><b>画面/口播/字幕</b><br>{safe_text(item.get('画面/口播/字幕'))}</p><dl class="detail-grid compact">{detail('停留作用',item.get('停留作用'))}{detail('信任作用',item.get('信任作用'))}{detail('成交作用',item.get('成交作用'))}{detail('可复刻点',item.get('可复刻点'),'good')}{detail('风险/缺口',item.get('风险/缺口'),'risk')}{detail('来源目录',item.get('来源目录'))}</dl></div></article>''')
    return f'<section id="segments"><div class="section-head"><p>02 · 内容推进</p><h2>逐段拆解时间轴</h2><span>{safe_text(note)}</span></div><div class="timeline">{"".join(items)}</div></section>'


def replication_section(sheet):
    records, _, note = sheet_records(sheet)
    cards = []
    fields = ["目标用户购买前状态","购买触发","信任/价值建立","主要异议与处理","交易加速器","现在购买理由","必须保留模块","可替换表现元素","失效条件","团队拍摄清单","可复用脚本骨架","下一轮单变量测试","成功指标","事实核验项"]
    for item in records:
        cards.append(f'''<article class="strategy-card"><div class="card-top"><div><span class="badge">{safe_text(item.get('素材编号'))}</span><h3>{safe_text(item.get('成交模型'))}</h3></div></div><dl class="detail-grid">{"".join(detail(field,item.get(field),'risk' if field in ('失效条件','事实核验项') else 'good' if field in ('必须保留模块','成功指标') else '') for field in fields)}</dl></article>''')
    return f'<section id="replication"><div class="section-head"><p>03 · 团队复刻</p><h2>复刻策略与测试计划</h2><span>{safe_text(note)}</span></div>{"".join(cards)}</section>'


def evidence_section(sheet):
    records, _, note = sheet_records(sheet)
    cards = []
    for item in records:
        row = item["__row_index"]
        image_data = next((data for (r, _), data in sheet["images"].items() if r == row), "")
        image_html = f'<button class="image-button" data-image="{image_data}" aria-label="查看 {safe_text(item.get("证据编号"))} 大图"><img src="{image_data}" alt="{safe_text(item.get("证据编号"))} 关键画面"></button>' if image_data else '<div class="image-missing">图片缺失</div>'
        cards.append(f'''<article class="evidence-card">{image_html}<div class="evidence-body"><div class="evidence-meta"><span class="badge">证据编号：{safe_text(item.get('证据编号'))}</span><b>时间戳（秒）：{safe_text(item.get('时间戳（秒）'))} · 证据角色：{safe_text(item.get('证据角色'))}</b></div><h3><small>支持的分析结论</small>{safe_text(item.get('支持的分析结论'))}</h3><p class="boundary"><b>证据边界</b><br>{safe_text(item.get('证据边界'))}</p><p class="source">素材编号：{safe_text(item.get('素材编号'))} · 原始文件名：{safe_text(item.get('原始文件名'))} · 来源目录：{safe_text(item.get('来源目录'))}</p></div></article>''')
    return f'<section id="evidence"><div class="section-head"><p>04 · 画面佐证</p><h2>关键证据图集</h2><span>{safe_text(note)}</span></div><div class="evidence-grid">{"".join(cards)}</div></section>'


def build_html(sheets, source_name):
    materials, images = sheet_stats(sheets)
    by_name = {sheet["name"]: sheet for sheet in sheets}
    sections = "".join(filter(None, [overview_section(by_name["素材总览"]) if "素材总览" in by_name else "", segments_section(by_name["逐段拆解"]) if "逐段拆解" in by_name else "", replication_section(by_name["复刻策略"]) if "复刻策略" in by_name else "", evidence_section(by_name["画面证据"]) if "画面证据" in by_name else ""]))
    data = json.dumps({"source": source_name, "sheets": len(sheets), "materials": materials, "images": images}, ensure_ascii=False).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    return f'''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="qianchuan-report-template" content="{HTML_TEMPLATE_VERSION}">
<title>千川视频成交模型拆解</title>
<style>
:root{{--navy:#17324d;--blue:#2f75b5;--ink:#1f2937;--muted:#64748b;--line:#dbe4ee;--paper:#f4f7fa;--risk:#fff1f2;--good:#effaf4;--warn:#fff8e7;--card:#fff}}
*{{box-sizing:border-box}} body{{margin:0;background:var(--paper);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif}}
.hero{{background:linear-gradient(135deg,#132c44,#1f4f78);color:#fff;padding:34px clamp(20px,4vw,64px) 28px}} .hero-inner{{max-width:1500px;margin:auto}}
.kicker{{margin:0 0 8px;color:#a9d4f6;font-size:13px;font-weight:700;letter-spacing:.12em}} h1{{margin:0;font-size:clamp(26px,3vw,42px)}}
.subtitle{{max-width:920px;margin:12px 0 22px;color:#d9e8f4;line-height:1.7}} .stats{{display:flex;flex-wrap:wrap;gap:10px}}
.stat{{min-width:130px;padding:12px 16px;border:1px solid #ffffff28;border-radius:12px;background:#ffffff10}} .stat b{{display:block;font-size:22px}} .stat span{{font-size:12px;color:#c9dbea}}
.nav-wrap{{position:sticky;top:0;z-index:20;background:#fffffff2;border-bottom:1px solid var(--line);backdrop-filter:blur(10px)}} nav{{max-width:1180px;margin:auto;padding:10px 20px;display:flex;gap:8px;overflow:auto}} nav a{{white-space:nowrap;text-decoration:none;border-radius:9px;padding:10px 16px;background:#e8eef5;color:#31506b;font-weight:700}}
main{{max-width:1180px;margin:0 auto 56px;padding:0 20px}} section{{padding-top:54px}} .section-head{{margin-bottom:18px}} .section-head p{{margin:0 0 5px;color:var(--blue);font-size:13px;font-weight:800;letter-spacing:.08em}} .section-head h2{{margin:0 0 8px;font-size:28px}} .section-head span{{display:block;color:var(--muted);line-height:1.6}}
.overview-card,.strategy-card,.timeline-card,.evidence-card{{background:#fff;border:1px solid var(--line);border-radius:16px;box-shadow:0 8px 24px #17324d0c}} .overview-card,.strategy-card{{padding:24px;margin-bottom:20px}} .card-top{{display:flex;justify-content:space-between;gap:20px;align-items:flex-start}} .card-top h3{{margin:10px 0 6px;font-size:23px;line-height:1.4}} .badge{{display:inline-block;background:#dcecf8;color:#1e5d8e;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:800}} .meta,.source{{color:var(--muted);font-size:12px}} .priority{{min-width:72px;text-align:center;background:#17324d;color:#fff;border-radius:12px;padding:11px;font-size:20px;font-weight:800}} .priority small{{display:block;color:#cbddea;font-size:10px}}
.hook,.chain{{margin-top:18px;padding:18px;border-radius:12px;background:#f0f6fb}} .hook p{{margin:8px 0 0;line-height:1.7}} .chain-flow{{display:flex;align-items:stretch;gap:8px;overflow:auto;margin-top:10px}} .chain-flow span{{position:relative;min-width:160px;padding:12px 28px 12px 14px;background:#fff;border:1px solid #cfe0ee;border-radius:10px;line-height:1.5}} .chain-flow span:not(:last-child)::after{{content:'→';position:absolute;right:-9px;top:50%;transform:translateY(-50%);z-index:2;color:var(--blue);font-weight:900}}
.detail-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin:18px 0 0}} .detail{{padding:14px;border:1px solid #e4ebf1;border-radius:11px;background:#fafcfe}} .detail.good{{background:var(--good)}} .detail.warn{{background:var(--warn)}} .detail.risk{{background:var(--risk)}} dt{{color:#50708b;font-size:12px;font-weight:800}} dd{{margin:6px 0 0;line-height:1.65}} .compact{{margin-top:12px}}
.timeline{{position:relative}} .timeline::before{{content:'';position:absolute;left:93px;top:0;bottom:0;width:2px;background:#bad2e5}} .timeline-item{{position:relative;display:grid;grid-template-columns:76px 1fr;gap:36px;margin-bottom:18px}} .time{{padding-top:18px;text-align:right}} .time b{{display:block;color:#1f5f90}} .time span{{display:block;margin-top:5px;color:var(--muted);font-size:11px}} .timeline-card{{position:relative;padding:20px}} .timeline-card::before{{content:'';position:absolute;left:-28px;top:24px;width:12px;height:12px;border:4px solid #fff;border-radius:50%;background:var(--blue);box-shadow:0 0 0 2px var(--blue)}} .timeline-title{{display:flex;justify-content:space-between;gap:12px}} .timeline-title h3{{margin:0}} .timeline-title span{{color:var(--blue);font-size:12px;font-weight:700}} .segment-meta{{color:var(--muted);font-size:12px}} .scene{{padding:14px;background:#f5f8fb;border-radius:10px;line-height:1.7}}
.evidence-grid{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}} .evidence-card{{overflow:hidden}} .image-button{{display:block;width:100%;height:330px;border:0;padding:0;background:#101820;cursor:zoom-in}} .image-button img{{width:100%;height:100%;object-fit:contain}} .evidence-body{{padding:18px}} .evidence-meta{{display:flex;align-items:center;gap:10px}} .evidence-body h3{{font-size:17px;line-height:1.55}} .evidence-body h3 small{{display:block;color:var(--blue);font-size:11px;margin-bottom:5px}} .boundary{{padding:12px;background:var(--risk);border-left:3px solid #df6b75;border-radius:6px;line-height:1.6}}
.modal{{position:fixed;inset:0;z-index:50;display:none;place-items:center;padding:24px;background:#08131eea}} .modal.open{{display:grid}} .modal img{{max-width:min(100%,1100px);max-height:90vh;border-radius:12px}} .modal button{{position:absolute;top:18px;right:22px;border:0;background:#fff;color:#132c44;border-radius:999px;width:42px;height:42px;font-size:24px;cursor:pointer}}
footer{{max-width:1500px;margin:0 auto 32px;padding:0 42px;color:var(--muted);font-size:12px}} code{{font-family:ui-monospace,SFMono-Regular,Menlo,monospace}}
@media(max-width:760px){{.detail-grid,.evidence-grid{{grid-template-columns:1fr}} .timeline::before{{left:7px}} .timeline-item{{grid-template-columns:1fr;gap:6px;padding-left:28px}} .time{{text-align:left;padding:0}} .timeline-card::before{{left:-27px;top:20px}} .card-top{{display:block}} .priority{{margin-top:12px;width:72px}}}}
@media print{{.nav-wrap{{display:none}} section{{break-before:page}} .overview-card,.strategy-card,.timeline-card,.evidence-card{{box-shadow:none}}}}
</style></head>
<body><header class="hero"><div class="hero-inner"><p class="kicker">QIANCHUAN VIDEO REVIEW</p><h1>千川视频成交模型拆解</h1><p class="subtitle">把素材结论、成交链、时间轴、复刻策略和画面证据整合成一份可连续阅读的分析报告。图片已嵌入本文件，断网也可打开。</p><div class="stats"><div class="stat"><b>{materials}</b><span>素材数量</span></div><div class="stat"><b>{len(sheets)}</b><span>分析模块</span></div><div class="stat"><b>{images}</b><span>证据图片</span></div></div></div></header>
<div class="nav-wrap"><nav><a href="#overview">核心结论</a><a href="#segments">逐段拆解</a><a href="#replication">复刻策略</a><a href="#evidence">画面证据</a></nav></div><main>{sections}</main>
<footer>来源文件：<code>{html.escape(source_name)}</code> · HTML Template v{HTML_TEMPLATE_VERSION} · 单文件离线报告 · 画面证据不能单独证明口播、经营事实或投放效果</footer>
<div class="modal" role="dialog" aria-modal="true"><button aria-label="关闭">×</button><img alt="证据图片大图"></div>
<script>const reportMeta={data};const m=document.querySelector('.modal'),mi=m.querySelector('img');document.querySelectorAll('.image-button').forEach(b=>b.onclick=()=>{{mi.src=b.dataset.image;m.classList.add('open')}});m.onclick=e=>{{if(e.target===m||e.target.tagName==='BUTTON')m.classList.remove('open')}};addEventListener('keydown',e=>{{if(e.key==='Escape')m.classList.remove('open')}});</script></body></html>'''


def main():
    parser = argparse.ArgumentParser(description="把千川拆解 Excel 转成内嵌图片的单文件 HTML")
    parser.add_argument("xlsx", help="输入 .xlsx 文件")
    parser.add_argument("output", nargs="?", help="输出 .html；默认与 Excel 同名")
    args = parser.parse_args()
    source = Path(args.xlsx).expanduser().resolve()
    if not source.is_file():
        raise SystemExit(f"Excel 不存在：{source}")
    output = Path(args.output).expanduser().resolve() if args.output else source.with_suffix(".html")
    sheets = parse_workbook(source)
    if not sheets:
        raise SystemExit("Excel 中没有可读取的工作表")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(build_html(sheets, source.name), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
