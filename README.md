# MangaEpubAutomation

面向漫画资源的自动化工具，提供两条流水线：

1. 目录流水线：图片超分 -> 分章/分卷 EPUB -> 合并大 EPUB  
2. EPUB 流水线：输入 EPUB（单本/多本/目录递归）-> 图片超分 -> 输出新 EPUB

同时提供 CLI（PowerShell）和 GUI（WPF + Wpf.Ui）。

## 你需要先准备的依赖（用户自行安装）

- CopyManga 下载器（用于准备输入目录）
  - https://github.com/lanyeeee/copymanga-downloader
- MangaJaNaiConverterGui（提供 python/backend/models）
  - https://github.com/the-database/MangaJaNaiConverterGui
- KCC / comic2ebook（用于章节 EPUB 打包）
  - https://github.com/ciromattia/kcc

说明：
- 本仓库不分发漫画内容、模型文件、KCC/MangaJaNai 本体。
- 请在 `manga_epub_automation.deps.json` 中填写本机依赖路径。

## 项目结构（核心文件）

```text
.
├─ Invoke-MangaEpubAutomation.ps1        # 目录流水线入口
├─ Invoke-EpubUpscalePipeline.ps1        # EPUB 流水线入口
├─ MergeEpubByOrder.py                   # 合并 EPUB 脚本
├─ manga_epub_automation.config.json     # 主配置（可调参数）
├─ manga_epub_automation.settings.json   # MangaJaNai 基础设置模板
├─ manga_epub_automation.deps.template.json
├─ gui/
│  └─ MangaEpubAutomation.Gui/
└─ logs/                                 # 运行产物
```

## 输入目录要求（目录流水线）

`-TitleRoot` 需指向单作品根目录，常见结构：

```text
<TitleRoot>/
  <SourceName>/
    元数据.json
    默認/
      <章节目录>/
        图片文件 + 章节元数据.json
    单行本/            # 可选
      <卷目录>/
        图片文件 + 章节元数据.json
```

关键规则：
- `单行本/` 不存在是允许的，会自动跳过卷处理。
- `默認/` 下若出现严格匹配“第xx卷/巻”的目录，会按卷处理。
- 其他额外分组可用 `-ExtraGroupHandlingMode` 控制（忽略/按话处理/按话并允许 merge）。

## 快速开始（目录流水线 CLI）

1. 复制依赖模板并填写路径：

```powershell
Copy-Item .\manga_epub_automation.deps.template.json .\manga_epub_automation.deps.json
```

2. 先预检：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -PlanOnly
```

3. 正式执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>"
```

常用阶段组合：

```powershell
# 仅超分
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipEpubPackaging -SkipMergedEpub

# 仅章节 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipMergedEpub

# 仅 merge 大 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipEpubPackaging
```

## EPUB 超分流水线（CLI）

入口：`Invoke-EpubUpscalePipeline.ps1`

输入模式（二选一）：
- `-InputEpubPath`：按文件（支持多本）
- `-InputEpubDirectory`：按目录递归

输出相关：
- `-OutputDirectory` 可选，不填时默认当前工作目录
- `-OutputSuffix` 默认 `-upscaled`
- 需要“无后缀”时，建议用 `-NoOutputSuffix`（比 `-OutputSuffix ""` 更稳定）

示例：

```powershell
# 单本
powershell -ExecutionPolicy Bypass -File .\Invoke-EpubUpscalePipeline.ps1 `
  -InputEpubPath "D:\Comics\Book01.epub" `
  -PlanOnly

# 多本（powershell -File 建议用 | 拼接）
powershell -ExecutionPolicy Bypass -File .\Invoke-EpubUpscalePipeline.ps1 `
  -InputEpubPath "D:\Comics\Book01.epub|D:\Comics\Book02.epub" `
  -OutputDirectory "D:\Comics\out" `
  -NoOutputSuffix
```

## GUI 使用

```powershell
dotnet restore .\MangaEpubAutomation.Gui.sln
dotnet run --project .\gui\MangaEpubAutomation.Gui\MangaEpubAutomation.Gui.csproj
```

重要说明：
- `Run` 页是目录流水线。
- `EPUB Pipeline` 页是 EPUB 流水线。
- 运行前建议先点 `Generate Plan` 看预检和计划。

## 合并顺序覆盖（可选）

默认文件：

```text
<TitleRoot>\<SourceName>-output\.manga_epub_automation_merge_order.json
```

格式：

```json
{
  "version": 1,
  "chapters": ["117话", "118话", "119话"]
}
```

`chapters` 的顺序就是 merge 优先顺序；未覆盖的章节会自动追加。

## 日志与运行产物

`logs/` 会生成：
- `run_plan_*.json` / `run_result_*.json`（目录流水线）
- `epub_run_plan_*.json` / `epub_run_result_*.json`（EPUB 流水线）
- `*_run_*.log`
- `latest_*.json` 快捷文件

## 临时文件行为（EPUB 流水线）

- 每本 EPUB 会在系统临时目录创建工作区（解包、超分中间结果、重打包）。
- 成功或失败都会在 `finally` 中自动清理该临时目录。
- 异常强退时可能残留，可手动清理系统临时目录。

## License

本仓库代码采用 MIT，见 [LICENSE](LICENSE)。
