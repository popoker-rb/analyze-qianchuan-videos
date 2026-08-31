# WorkBuddy 使用说明

## 放置方式

将整个 `analyze-qianchuan-videos` 文件夹复制到 WorkBuddy 可读取的 Skills 目录，保持 `SKILL.md`、`references/` 和 `scripts/` 的相对位置不变。不同 WorkBuddy 安装的 Skill 根目录可能不同，先在目标电脑确认，不要猜固定路径。

## 首次诊断

在 Skill 根目录运行：

```text
python scripts/diagnose.py
```

结果中：

- `video_review_ready=true`：可自动盘点视频并生成联系表。
- 缺 `ffmpeg` 或 `ffprobe`：不要假装已观看视频；报告缺失能力，请用户或管理员配置。
- 缺 `node`：不能运行随附 Excel 构建器，但如果 WorkBuddy 有原生 Excel/电子表格能力，可按输出契约直接生成。

不自动联网安装软件，不修改系统环境变量。

## 推荐调用话术

```text
请使用 analyze-qianchuan-videos Skill，逐条分析这个文件夹里的视频素材。识别每条的成交模型、跑量条件、优点、不足、合规风险和团队复刻方法，完成后做横向总结，同步输出Excel和一个内嵌图片的整合型单文件HTML。
```

如果用户要求先看一条：

```text
先只分析自然排序第1条，同步输出Excel和整合型单文件HTML让我确认；确认后再继续剩余视频，不要提前批量分析。
```

## Excel 路由

优先级：

1. WorkBuddy 原生 Excel/电子表格能力；
2. 环境已经能解析 `@oai/artifact-tool` 时使用 `scripts/build_report.mjs`；
3. 两者都不可用时停止并报告，不能用 CSV、Markdown 或 HTML 冒充用户要求的 Excel。

Excel 生成后运行 `python scripts/xlsx_to_single_html.py <xlsx>`。HTML 必须整合为分析报告，不照搬 Sheet 表格；转换器直接读取 Excel 内嵌图片，不依赖外部图片目录。

## 安全边界

- 视频只读，不移动、不重命名、不上传。
- 中间联系表和证据帧放项目 `work/`，正式 Excel 放 `deliverables/`。
- 只选能佐证分析的关键帧，通常每条3–5张；不得为了让报告好看而堆图。
- 不自动发送报告；由用户确认后自行分发。
- 用户提供“跑量”描述不等于已提供投放数据，报告仍需注明证据边界。
