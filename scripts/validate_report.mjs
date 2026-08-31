import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const file = process.argv[2];
if (!file) throw new Error("Usage: node scripts/validate_report.mjs <output.xlsx>");
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const required = ["素材总览", "逐段拆解", "复刻策略", "画面证据"];
const names = wb.worksheets.items.map(s => s.name);
const missingSheets = required.filter(n => !names.includes(n));
const requiredHeaders = {
  "素材总览": ["素材编号", "原始文件名", "成交模型", "模型状态", "一句话成交链", "跑量逻辑", "合规风险"],
  "逐段拆解": ["素材编号", "起始时间", "结束时间", "画面/口播/字幕", "成交作用"],
  "复刻策略": ["素材编号", "成交模型", "购买触发", "主要异议与处理", "交易加速器", "失效条件", "下一轮单变量测试"],
  "画面证据": ["证据编号", "素材编号", "时间戳（秒）", "证据角色", "支持的分析结论", "证据边界", "关键画面"]
};
const missingHeaders = [];
for (const name of required.filter(n => names.includes(n))) {
  const sheet = wb.worksheets.getItem(name);
  const used = sheet.getUsedRange();
  const width = Math.max(used?.columnCount || 1, 1);
  const row = sheet.getRangeByIndexes(3, 0, 1, width).values[0].map(String);
  for (const header of requiredHeaders[name]) if (!row.includes(header)) missingHeaders.push(`${name}:${header}`);
}
const scan = JSON.parse((await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 } })).ndjson);
const result = { file, sheets: names, missingSheets, missingHeaders, formulaErrorScan: scan };
console.log(JSON.stringify(result, null, 2));
if (missingSheets.length || missingHeaders.length) process.exitCode = 2;
