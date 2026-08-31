#!/usr/bin/env python3
"""验证整合型单文件 HTML 与源 Excel 一致且未依赖外部资源。"""

import argparse
import html
import importlib.util
import json
import re
from pathlib import Path


def load_converter():
    path = Path(__file__).with_name("xlsx_to_single_html.py")
    spec = importlib.util.spec_from_file_location("xlsx_to_single_html", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate(xlsx_path, html_path):
    converter = load_converter()
    sheets = converter.parse_workbook(xlsx_path)
    source = Path(html_path).read_text(encoding="utf-8")
    required_sections = ["overview", "segments", "replication", "evidence"]
    missing_sections = [name for name in required_sections if f'id="{name}"' not in source]
    missing_cells = []
    for sheet in sheets:
        rows_to_check = sheet["rows"][:3] + sheet["rows"][4:]
        for row in rows_to_check:
            for value in row:
                expected = html.escape(str(value)).replace("\n", "<br>") if value not in (None, "") else ""
                if expected and expected not in source:
                    missing_cells.append({"sheet": sheet["name"], "value": str(value)[:120]})
    excel_images = sum(len(sheet["images"]) for sheet in sheets)
    html_images = source.count('class="image-button"')
    external_refs = re.findall(r'<(?:link\b[^>]*\bhref|script\b[^>]*\bsrc|img\b[^>]*\bsrc)="(?!data:)([^"#]+)"', source, re.I)
    result = {
        "templateVersion": converter.HTML_TEMPLATE_VERSION,
        "templateMarkerPresent": f'content="{converter.HTML_TEMPLATE_VERSION}"' in source,
        "requiredSections": required_sections,
        "missingSections": missing_sections,
        "spreadsheetTables": len(re.findall(r"<table\b", source, re.I)),
        "excelImageCount": excel_images,
        "htmlImageCount": html_images,
        "imageCountMatches": excel_images == html_images,
        "externalResourceReferences": external_refs,
        "missingExcelCellCount": len(missing_cells),
        "missingExcelCells": missing_cells[:20],
    }
    result["passed"] = all([
        result["templateMarkerPresent"],
        not missing_sections,
        result["spreadsheetTables"] == 0,
        result["imageCountMatches"],
        not external_refs,
        not missing_cells,
    ])
    return result


def main():
    parser = argparse.ArgumentParser(description="验证整合型单文件 HTML")
    parser.add_argument("xlsx")
    parser.add_argument("html")
    args = parser.parse_args()
    result = validate(Path(args.xlsx).expanduser().resolve(), Path(args.html).expanduser().resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result["passed"]:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
