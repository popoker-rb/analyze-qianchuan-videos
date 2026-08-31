#!/usr/bin/env python3
import json
import shutil

result = {
    "python": shutil.which("python") or shutil.which("python3"),
    "node": shutil.which("node"),
    "ffmpeg": shutil.which("ffmpeg"),
    "ffprobe": shutil.which("ffprobe"),
}
result["video_review_ready"] = bool(result["ffmpeg"] and result["ffprobe"])
result["builder_runtime_present"] = bool(result["node"])
print(json.dumps(result, ensure_ascii=False, indent=2))
