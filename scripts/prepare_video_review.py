#!/usr/bin/env python3
import argparse
import json
import shutil
import subprocess
from pathlib import Path

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".avi", ".mkv", ".webm"}

def run(cmd):
    return subprocess.run(cmd, check=True, capture_output=True, text=True).stdout

def main():
    parser = argparse.ArgumentParser(description="盘点视频并生成抽帧联系表")
    parser.add_argument("--input", required=True, help="视频文件或目录")
    parser.add_argument("--work", required=True, help="中间文件目录")
    parser.add_argument("--interval", type=float, default=2.0, help="抽帧间隔秒数")
    args = parser.parse_args()
    if not shutil.which("ffprobe") or not shutil.which("ffmpeg"):
        raise SystemExit("缺少 ffprobe 或 ffmpeg；请先运行 scripts/diagnose.py")
    source = Path(args.input).expanduser().resolve()
    files = [source] if source.is_file() else sorted(p for p in source.iterdir() if p.suffix.lower() in VIDEO_EXTS)
    if not files:
        raise SystemExit("没有找到支持的视频文件")
    work = Path(args.work).expanduser().resolve()
    contacts = work / "contact_sheets"
    contacts.mkdir(parents=True, exist_ok=True)
    inventory = []
    for index, video in enumerate(files, 1):
        probe = json.loads(run(["ffprobe", "-v", "error", "-show_entries", "format=duration,size:stream=codec_type,width,height,r_frame_rate", "-of", "json", str(video)]))
        vstream = next((s for s in probe.get("streams", []) if s.get("codec_type") == "video"), {})
        audio = any(s.get("codec_type") == "audio" for s in probe.get("streams", []))
        duration = float(probe.get("format", {}).get("duration") or 0)
        safe = f"{index:02d}_{video.stem[:60]}"
        contact_pattern = contacts / f"{safe}_%03d.jpg"
        fps = 1 / max(args.interval, 0.25)
        run(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", str(video), "-vf", f"fps={fps},scale=216:-1,tile=5x6:nb_frames=30:padding=4:margin=4", str(contact_pattern)])
        pages = sorted(str(p) for p in contacts.glob(f"{safe}_*.jpg"))
        inventory.append({
            "materialNo": f"{index:02d}", "fileName": video.name, "path": str(video),
            "duration": round(duration, 3), "width": vstream.get("width"), "height": vstream.get("height"),
            "frameRate": vstream.get("r_frame_rate"), "hasAudio": audio, "contactSheets": pages
        })
    out = work / "inventory.json"
    out.write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding="utf-8")
    print(out)

if __name__ == "__main__":
    main()
