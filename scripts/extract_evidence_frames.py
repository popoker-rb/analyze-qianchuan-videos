#!/usr/bin/env python3
import argparse
import shutil
import subprocess
from pathlib import Path

def main():
    parser = argparse.ArgumentParser(description="按时间戳提取Excel画面证据帧")
    parser.add_argument("--video", required=True)
    parser.add_argument("--timestamps", required=True, help="逗号分隔秒数，例如 0.5,6,14.2")
    parser.add_argument("--output", required=True)
    parser.add_argument("--material-no", required=True)
    args = parser.parse_args()
    if not shutil.which("ffmpeg"):
        raise SystemExit("缺少 ffmpeg；请先运行 scripts/diagnose.py")
    video = Path(args.video).expanduser().resolve()
    output = Path(args.output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    for index, raw in enumerate(args.timestamps.split(","), 1):
        timestamp = float(raw.strip())
        target = output / f"{args.material_no}-E{index:02d}_{timestamp:.1f}s.jpg"
        subprocess.run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-ss", str(timestamp), "-i", str(video), "-frames:v", "1", "-q:v", "2", str(target)], check=True)
        print(target)

if __name__ == "__main__":
    main()
