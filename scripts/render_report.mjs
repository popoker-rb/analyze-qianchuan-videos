import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const file = process.argv[2];
const outputDir = process.argv[3] || path.join(path.dirname(file || "."), "previews");
if (!file) throw new Error("Usage: node scripts/render_report.mjs <output.xlsx> [preview-dir]");
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
await fs.mkdir(outputDir, { recursive: true });
for (const sheet of wb.worksheets.items) {
  const blob = await wb.render({ sheetName: sheet.name, autoCrop: "all", scale: 1 });
  await fs.writeFile(path.join(outputDir, `${sheet.name}.png`), new Uint8Array(await blob.arrayBuffer()));
}
console.log(outputDir);
