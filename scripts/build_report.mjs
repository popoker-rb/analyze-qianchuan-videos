import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const input = process.argv[2];
if (!input) throw new Error("Usage: node scripts/build_report.mjs <report-data.json>");
const cfg = JSON.parse(await fs.readFile(input, "utf8"));
const base = path.dirname(path.resolve(input));
const output = path.resolve(base, cfg.outputPath || "deliverables/视频素材成交模型拆解.xlsx");
const wb = Workbook.create();

const specs = [
  { name: "素材总览", rows: cfg.overview || [], columns: [
    ["materialNo","素材编号",10],["fileName","原始文件名",26],["duration","时长（秒）",11,"0.0"],["format","画幅",15],
    ["product","产品/品类",15],["hook","核心钩子",28],["conversionModel","成交模型",24],["modelStatus","模型状态",13],
    ["conversionChain","一句话成交链",42],["evidenceRefs","画面证据编号",18],["scalingLogic","跑量逻辑",42],["strength","最大优点",28],["weakness","主要不足",28],
    ["complianceRisk","合规风险",32],["confidence","模型置信度",12],["poolRole","建议素材池角色",16],["priority","优先级",10]
  ]},
  { name: "逐段拆解", rows: cfg.segments || [], columns: [
    ["materialNo","素材编号",10],["fileName","原始文件名",24],["start","起始时间",10,"0.0"],["end","结束时间",10,"0.0"],
    ["content","画面/口播/字幕",42],["evidenceRefs","画面证据编号",18],["function","内容功能",18],["retentionRole","停留作用",28],["trustRole","信任作用",28],
    ["conversionRole","成交作用",30],["replicable","可复刻点",32],["risk","风险/缺口",30]
  ]},
  { name: "复刻策略", rows: cfg.replication || [], columns: [
    ["materialNo","素材编号",10],["conversionModel","成交模型",24],["prePurchaseState","目标用户购买前状态",34],
    ["purchaseTrigger","购买触发",28],["trustValue","信任/价值建立",34],["objectionHandling","主要异议与处理",34],
    ["accelerator","交易加速器",28],["buyNowReason","现在购买理由",28],["mustKeep","必须保留模块",36],
    ["replaceable","可替换表现元素",30],["failureConditions","失效条件",32],["shotList","团队拍摄清单",38],
    ["scriptSkeleton","可复用脚本骨架",42],["nextTest","下一轮单变量测试",32],["successMetrics","成功指标",28],["factChecks","事实核验项",32]
  ]}
];

for (const spec of specs) addSheet(spec);
await addEvidenceSheet(cfg.evidence || []);
await fs.mkdir(path.dirname(output), { recursive: true });
const blob = await SpreadsheetFile.exportXlsx(wb);
await blob.save(output);
const htmlOutput = output.replace(/\.xlsx$/i, ".html");
const converter = path.join(path.dirname(fileURLToPath(import.meta.url)), "xlsx_to_single_html.py");
const htmlValidator = path.join(path.dirname(fileURLToPath(import.meta.url)), "validate_single_html.py");
const pythonCandidates = [process.env.PYTHON, "python", "python3"].filter(Boolean);
let converted = false;
let lastError = "";
for (const python of pythonCandidates) {
  const result = spawnSync(python, [converter, output, htmlOutput], { encoding: "utf8" });
  if (!result.error && result.status === 0) { converted = true; break; }
  lastError = result.error?.message || result.stderr || `exit ${result.status}`;
}
if (!converted) throw new Error(`Excel 已生成，但单文件 HTML 生成失败：${lastError}`);
let validated = false;
for (const python of pythonCandidates) {
  const result = spawnSync(python, [htmlValidator, output, htmlOutput], { encoding: "utf8" });
  if (!result.error && result.status === 0) { validated = true; break; }
  lastError = result.error?.message || result.stderr || result.stdout || `exit ${result.status}`;
}
if (!validated) throw new Error(`单文件 HTML 自动验收失败：${lastError}`);
console.log(JSON.stringify({ xlsx: output, html: htmlOutput }, null, 2));

function addSheet(spec) {
  const s = wb.worksheets.add(spec.name);
  s.showGridLines = false;
  const cols = spec.columns;
  const last = colName(cols.length);
  s.getRange(`A1:${last}1`).merge();
  s.getRange("A1").values = [[cfg.reportTitle || "视频素材成交模型拆解"]];
  s.getRange(`A1:${last}1`).format = { fill: "#17324D", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 34 };
  s.getRange(`A2:${last}2`).merge();
  s.getRange("A2").values = [[`来源：${cfg.sourcePath || "未提供"}｜口径：${cfg.scopeNote || "未提供"}`]];
  s.getRange(`A2:${last}2`).format = { fill: "#F2F5F8", font: { color: "#334155", size: 10 }, wrapText: true, rowHeight: 32 };
  s.getRange(`A4:${last}4`).values = [cols.map(c => c[1])];
  s.getRange(`A4:${last}4`).format = { fill: "#2F75B5", font: { bold: true, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "center", rowHeight: 42, borders: { preset: "outside", style: "thin", color: "#D7E0EA" } };
  if (spec.rows.length) {
    s.getRange(`A5:${last}${4 + spec.rows.length}`).values = spec.rows.map(row => cols.map(c => row[c[0]] ?? null));
    const body = s.getRange(`A5:${last}${4 + spec.rows.length}`);
    body.format = { font: { color: "#243447", size: 10 }, wrapText: true, verticalAlignment: "top", borders: { insideHorizontal: { style: "thin", color: "#E6EBF0" } }, rowHeight: spec.name === "逐段拆解" ? 72 : 96 };
    cols.forEach((c, i) => {
      s.getRangeByIndexes(0, i, 1, 1).format.columnWidth = c[2];
      if (c[3]) s.getRangeByIndexes(4, i, spec.rows.length, 1).format.numberFormat = c[3];
    });
    const statusIndex = cols.findIndex(c => c[0] === "modelStatus");
    if (statusIndex >= 0) {
      const r = s.getRangeByIndexes(4, statusIndex, spec.rows.length, 1);
      r.conditionalFormats.add("containsText", { text: "推定", format: { fill: "#FFF4CC" } });
      r.conditionalFormats.add("containsText", { text: "模型级验证", format: { fill: "#DDF3E4" } });
    }
  } else cols.forEach((c, i) => { s.getRangeByIndexes(0, i, 1, 1).format.columnWidth = c[2]; });
  s.freezePanes.freezeRows(4);
  s.freezePanes.freezeColumns(spec.name === "素材总览" ? 2 : 1);
}

async function addEvidenceSheet(rows) {
  const s = wb.worksheets.add("画面证据");
  s.showGridLines = false;
  const cols = [
    ["evidenceId","证据编号",12],["materialNo","素材编号",10],["fileName","原始文件名",26],["timestamp","时间戳（秒）",12],
    ["role","证据角色",18],["finding","支持的分析结论",42],["boundary","证据边界",34],["imagePath","关键画面",24]
  ];
  const last = colName(cols.length);
  s.getRange(`A1:${last}1`).merge();
  s.getRange("A1").values = [[cfg.reportTitle || "视频素材成交模型拆解"]];
  s.getRange(`A1:${last}1`).format = { fill: "#17324D", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 34 };
  s.getRange(`A2:${last}2`).merge();
  s.getRange("A2").values = [["关键帧只证明对应时间点的可见画面；不能单独证明口播、经营事实或投放效果。"]];
  s.getRange(`A2:${last}2`).format = { fill: "#F2F5F8", font: { color: "#334155", size: 10 }, wrapText: true, rowHeight: 32 };
  s.getRange(`A4:${last}4`).values = [cols.map(c => c[1])];
  s.getRange(`A4:${last}4`).format = { fill: "#2F75B5", font: { bold: true, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "center", rowHeight: 42 };
  cols.forEach((c, i) => { s.getRangeByIndexes(0, i, 1, 1).format.columnWidth = c[2]; });
  if (rows.length) {
    const values = rows.map(row => cols.map(c => c[0] === "imagePath" ? "待嵌入" : (row[c[0]] ?? null)));
    s.getRange(`A5:${last}${4 + rows.length}`).values = values;
    s.getRange(`A5:${last}${4 + rows.length}`).format = { font: { color: "#243447", size: 10 }, wrapText: true, verticalAlignment: "top", borders: { insideHorizontal: { style: "thin", color: "#E6EBF0" } }, rowHeight: 108 };
    s.getRange(`D5:D${4 + rows.length}`).format.numberFormat = "0.0";
    for (let i = 0; i < rows.length; i++) {
      const rawPath = rows[i].imagePath;
      if (!rawPath) { s.getCell(4 + i, 7).values = [["图片文件缺失"]]; continue; }
      const imagePath = path.isAbsolute(rawPath) ? rawPath : path.resolve(base, rawPath);
      try {
        const bytes = await fs.readFile(imagePath);
        const { width, height } = imageDimensions(bytes);
        const scale = Math.min(150 / width, 96 / height);
        const widthPx = Math.max(1, Math.round(width * scale));
        const heightPx = Math.max(1, Math.round(height * scale));
        const mime = path.extname(imagePath).toLowerCase() === ".png" ? "image/png" : "image/jpeg";
        const dataUrl = `data:${mime};base64,${bytes.toString("base64")}`;
        s.images.add({ dataUrl, anchor: { from: { row: 4 + i, col: 7, rowOffsetPx: 5, colOffsetPx: 5 }, extent: { widthPx, heightPx } } });
        s.getCell(4 + i, 7).values = [[""]];
      } catch (error) {
        s.getCell(4 + i, 7).values = [[`图片文件缺失：${rawPath}`]];
      }
    }
  }
  s.freezePanes.freezeRows(4);
  s.freezePanes.freezeColumns(2);
}

function imageDimensions(bytes) {
  if (bytes[0] === 0x89 && bytes.toString("ascii", 1, 4) === "PNG") return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) };
  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) { offset++; continue; }
      const marker = bytes[offset + 1];
      const length = bytes.readUInt16BE(offset + 2);
      if (marker >= 0xc0 && marker <= 0xc3) return { height: bytes.readUInt16BE(offset + 5), width: bytes.readUInt16BE(offset + 7) };
      offset += 2 + length;
    }
  }
  return { width: 160, height: 90 };
}

function colName(n) { let out = ""; while (n) { n--; out = String.fromCharCode(65 + n % 26) + out; n = Math.floor(n / 26); } return out; }
