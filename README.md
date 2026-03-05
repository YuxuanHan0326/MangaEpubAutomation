# MangaEpubAutomation

用于漫画图片处理的 CLI 流水线：
1. 超分（MangaJaNai backend）
2. 分章 EPUB 打包（KCC）
3. 默認分组合并大 EPUB（按 `order` 排序）

## 上游与第三方依赖

1. `copymanga-downloader`（下载并分类）
- https://github.com/misaka10843/copymanga-downloader

2. `MangaJaNaiConverterGui`（提供 Python/backend/models 运行环境）
- https://github.com/the-database/MangaJaNaiConverterGui

3. `KCC`（分章打包 EPUB）
- https://github.com/ciromattia/kcc

注意：
- `MangaJaNaiConverterGui` 和 `KCC` 都应由用户自行从官方仓库下载、安装、更新。
- 本项目不分发第三方二进制、模型文件和用户本机运行缓存。

## 许可证

本仓库代码采用 `MIT` 许可证，见 `LICENSE`。

## 推荐工作流

1. 用 `copymanga-downloader` 下载漫画（通常会产出本项目所需分组结构）。
2. 配置 `manga_epub_automation.deps.json`（第三方路径）。
3. 可选：配置 `manga_epub_automation.config.json`（MangaJaNai/KCC 运行参数默认值）。
4. 运行 `Invoke-MangaEpubAutomation.ps1 -TitleRoot <作品根目录>`。
5. 获取输出：
- `<SourceName>-upscaled`
- `<SourceName>-output`

## 输入目录结构（期望）

`-TitleRoot` 指向单作品根目录。

```text
<TitleRoot>/
  <SourceName>/
    元数据.json
    默認/
      <章节目录A>/
        *.jpg|*.jpeg|*.png|*.webp|*.avif|*.bmp
        章节元数据.json
      <章节目录B>/
    单行本/
      <卷目录A>/
        *.jpg|*.jpeg|*.png|*.webp|*.avif|*.bmp
        章节元数据.json
      <卷目录B>/
```

## 输出目录结构

```text
<TitleRoot>/
  <SourceName>-upscaled/
    默認/
      <章节目录A>/
        *-upscaled.webp|png|jpeg|avif
      <章节目录A>.epub
    单行本/
      <卷目录A>/
        *-upscaled.webp|png|jpeg|avif

  <SourceName>-output/
    <单行本epub>.epub
    <作品名 - 话范围>.epub
    .manga_epub_automation_merge_manifest.json
    .manga_epub_automation_merge_order.json
```

## 依赖配置

本地运行时：
1. 复制 `manga_epub_automation.deps.template.json` 为 `manga_epub_automation.deps.json`
2. 填写真实路径（Python/backend/models/KCC/merge_script）

示例：

```json
{
  "version": 1,
  "paths": {
    "python_exe": "<path-to-python.exe>",
    "backend_script": "<path-to-run_upscale.py>",
    "models_dir": "<path-to-models-dir>",
    "kcc_exe": "<path-to-kcc-exe>",
    "merge_script": "<path-to-MergeEpubByOrder.py>"
  },
  "progress": {
    "enabled": true,
    "refresh_seconds": 1,
    "eta_min_samples": 8,
    "noninteractive_log_interval_seconds": 10
  }
}
```

## 全局参数配置（manga_epub_automation.config.json）

这份文件用于统一管理原先脚本里 hardcoded 的 MangaJaNai/KCC/Merge 参数默认值。  
命令行参数仍可覆盖其中一部分（例如 `-UpscaleFactor/-OutputFormat/-LossyQuality`）。

默认文件名：
- `manga_epub_automation.config.json`

初始化默认模板：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -InitConfig
```

示例：

```json
{
  "version": 2,
  "pipeline": {
    "SourceDirSuffixToSkip": "-upscaled",
    "OutputDirSuffix": "-upscaled",
    "EpubOutputDirSuffix": "-output",
    "OutputFilenamePattern": "%filename%-upscaled"
  },
  "manga": {
    "SelectedTabIndex": 1,
    "OverwriteExistingFiles": false,
    "ModeScaleSelected": true,
    "ModeWidthSelected": false,
    "ModeHeightSelected": false,
    "ModeFitToDisplaySelected": false,
    "UpscaleScaleFactor": 2,
    "OutputFormat": "webp",
    "LossyCompressionQuality": 80,
    "WorkflowOverrides": {}
  },
  "kcc": {
    "CliOptions": {
      "DeviceProfile": "KS",
      "OutputFormat": "EPUB",
      "NoKepub": true,
      "DisableProcessing": true,
      "ForceColor": true,
      "AdditionalArgs": []
    },
    "UnicodeStaging": {
      "Enabled": true,
      "StagePrefix": "manga_epub_automation_kcc_stage_",
      "SafeTitleFallback": "manga_epub_automation_book",
      "SafeAuthorFallback": "pipeline"
    },
    "MetadataRewrite": {
      "Enabled": true
    },
    "BaseArgs": []
  },
  "merge": {
    "Language": "zh",
    "DescriptionHeader": "选集 包含:",
    "IncludeOrderInDescription": true,
    "MetadataContributor": "manga-epub-automation-merge"
  }
}
```

`kcc.CliOptions` 字段映射：
1. Main:
`DeviceProfile` -> `-p`
`OutputFormat` -> `-f`
`MangaStyle` -> `-m`
`HighQualityMagnification` -> `-q`
`TwoPanelView` -> `-2`
`WebtoonMode` -> `-w`
`TargetSizeMB` -> `--ts`

2. Processing:
`DisableProcessing` -> `-n`
`LegacyPdfExtract` -> `--pdfextract`
`UpscaleSmallImages` -> `-u`
`StretchToResolution` -> `-s`
`SplitterMode` -> `-r`
`Gamma` -> `-g`
`CroppingMode` -> `-c`
`CroppingPower` -> `--cp`
`PreserveMarginPercent` -> `--preservemargin`
`CroppingMinimumRatio` -> `--cm`
`InterPanelCropMode` -> `--ipc`
`ForceBlackBorders` -> `--blackborders`
`ForceWhiteBorders` -> `--whiteborders`
`ForceColor` -> `--forcecolor`
`ForcePng` -> `--forcepng`
`MozJpeg` -> `--mozjpeg`
`JpegQuality` -> `--jpeg-quality`
`MaximizeStrips` -> `--maximizestrips`
`DeleteSourceAfterPack` -> `-d`（危险操作，默认 false）

3. Output:
`NoKepub` -> `--nokepub`
`BatchSplitMode` -> `-b`
`MetadataTitleMode` -> `--metadatatitle`
`SpreadShift` -> `--spreadshift`
`NoRotateSpreads` -> `--norotate`
`RotateFirstSpread` -> `--rotatefirst`
`AutoLevel` -> `--autolevel`
`DisableAutoContrast` -> `--noautocontrast`
`ColorAutoContrast` -> `--colorautocontrast`
`FileFusion` -> `--filefusion`
`EraseRainbow` -> `--eraserainbow`

4. Custom profile:
`CustomWidth` -> `--customwidth`
`CustomHeight` -> `--customheight`

5. 额外参数：
`AdditionalArgs` -> 原样追加任意 KCC 参数（用于新版本未内置字段）

`manga.WorkflowOverrides` 用法：
1. 这是对 GUI workflow JSON 字段的“原名覆盖”入口
2. 用于补充脚本未显式支持但后端已支持的参数
3. 示例：`"WorkflowOverrides": { "TileSize": 512, "ModelFilePath": "xxx" }`

`merge` 配置含义：
1. `Language`：合并大 EPUB 的 `dc:language`
2. `DescriptionHeader`：描述首行
3. `IncludeOrderInDescription`：描述中是否写入 `order`
4. `MetadataContributor`：合并大 EPUB 的 `dc:contributor`

兼容说明：
1. 旧字段 `kcc.BaseArgs` 仍支持
2. 当 `kcc.BaseArgs` 非空时优先使用它（兼容旧配置）
3. 推荐新配置使用 `kcc.CliOptions`

## Merge 显式顺序文件

默认路径：
- `<TitleRoot>\<SourceName>-output\.manga_epub_automation_merge_order.json`

格式：

```json
{
  "version": 1,
  "chapters": [
    "117话",
    "118话",
    "119话"
  ]
}
```

规则：
1. `chapters` 顺序即 merge 优先顺序
2. 章节名精确匹配（trim 后）
3. 重复项仅首个生效，后续 `WARN`
4. 不存在章节会 `WARN` 并忽略
5. 未覆盖章节自动追加末尾（按自动排序）
6. 文件非法（JSON 错误/缺字段）会 `ERROR` 阻断

导出模板：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 `
  -TitleRoot "<TitleRoot>" `
  -DumpMergeOrderTemplate
```

## 常用命令

```powershell
# 帮助
Get-Help .\Invoke-MangaEpubAutomation.ps1 -Full

# 初始化配置模板
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -InitConfig
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -InitDepsConfig

# 全流程
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>"

# 仅计划预检
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -PlanOnly

# 仅超分
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipEpubPackaging -SkipMergedEpub

# 仅分章 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipMergedEpub

# 仅合并大 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipEpubPackaging
```

## 日志与机器接口

`logs/` 会生成：
- `run_plan_*.json`
- `run_result_*.json`
- `manga_epub_automation_run_*.log`

可用于后续 GUI 接入：
- `run_plan_*.json`：计划页
- `run_result_*.json`：结果页

## GUI 接入建议（新增）

启用事件流：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 `
  -TitleRoot "<TitleRoot>" `
  -GuiMode
```

`-GuiMode` 开启后，脚本会额外输出前缀为 `PIPELINE_EVENT:` 的 JSON 行，建议 GUI 直接按行订阅解析。

已提供的关键事件：
1. `plan_ready`：计划文件已生成
2. `preflight_summary`：预检统计（errors/warnings/infos + issues）
3. `confirmation_required` / `confirmation_response`：交互确认门禁
4. `stage`：阶段生命周期（`upscale|epub|merge` 的 `start|end|skip`）
5. `upscale_progress`：超分进度快照（percent/done/total/ETA/rate）
6. `epub_progress`：分章打包进度快照
7. `merge_preview`：merge 最终顺序预览（章节数组）
8. `run_result`：最终状态

稳定文件接口（便于 GUI 轮询）：
1. `logs/latest_run_plan.json`
2. `logs/latest_run_result.json`

说明：
1. `latest_*` 每次运行会覆盖为最近一次结果
2. GUI 不建议解析控制台普通文本，应优先读取 `PIPELINE_EVENT` 和 `latest_*` JSON


