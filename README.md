# MangaEpubAutomation

面向漫画图片目录的自动化流水线：

1. 图片超分（MangaJaNai backend）
2. 单章/单卷 EPUB 打包（KCC）
3. 默认分组章节合并为总集 EPUB（按 `order` 排序）

仓库同时提供：

- CLI 入口（PowerShell）
- GUI 前端（WPF + Wpf.Ui）

## 功能特性

- 增量处理：已存在输出自动跳过
- 可拆分阶段：可单跑超分/打包/合并
- 合并前预览：支持章节顺序预览与确认
- 显式顺序覆盖：可用 `merge_order.json` 手动指定合并顺序
- 计划与结果产物：`run_plan_*.json` / `run_result_*.json`

## 第三方依赖（用户自行安装）

- CopyManga 下载器（用于准备输入目录）
  - https://github.com/misaka10843/copymanga-downloader
- MangaJaNaiConverterGui（提供 python/backend/models）
  - https://github.com/the-database/MangaJaNaiConverterGui
- KCC / comic2ebook（用于章节 EPUB 打包）
  - https://github.com/ciromattia/kcc

说明：

- 本仓库不分发漫画内容、模型文件和上述依赖本体。
- `kcc` 可执行文件需用户自行下载并在依赖配置中填写路径。

## 仓库结构

```text
.
├─ Invoke-MangaEpubAutomation.ps1
├─ Invoke-MangaEpubAutomation.cmd
├─ MergeEpubByOrder.py
├─ manga_epub_automation.config.json
├─ manga_epub_automation.settings.json
├─ manga_epub_automation.deps.template.json
├─ gui/
│  └─ MangaEpubAutomation.Gui/
└─ logs/  # 运行时生成（默认忽略）
```

## 输入目录约定

`-TitleRoot` 指向单作品根目录，期望结构：

```text
<TitleRoot>/
  <SourceName>/
    元数据.json
    默認/
      <章节目录>/
        *.jpg|*.jpeg|*.png|*.webp|*.avif|*.bmp
        章节元数据.json
    单行本/
      <卷目录>/
        *.jpg|*.jpeg|*.png|*.webp|*.avif|*.bmp
        章节元数据.json
```

说明：

- `单行本/` 目录是可选的。若不存在会自动跳过单行本阶段，不影响默认章节流程。
- 若 `默認/` 下混有“数字+卷/巻”目录（例如 `12卷`、`第03巻`），会自动按单行本规则打包并输出到 `<SourceName>-output/`。
- 若存在除 `默認/`、`单行本/` 之外的其它一级分组目录，会按“话”规则进行章节 EPUB 打包（仅前两阶段），默认不参与总集 merge。
- 卷判定使用“严格模式”：仅当目录名完整匹配 `第?数字(可小数)+卷/巻`（例如 `第01卷`）才判卷；如 `连载版01`、`01卷附赠漫画`、`第3卷发售纪念` 会按“话/特别篇”处理。

额外分组处理模式（`-ExtraGroupHandlingMode`）：

- `ignore`（默认）：忽略 `默認/单行本` 之外的一级目录
- `chapter_no_merge`：参照“话”处理（打包章节 EPUB），但不参与 merge
- `chapter_merge`：参照“话”处理并允许 merge；输出名为 `作品名 - 文件夹名.epub`

## 输出目录约定

```text
<TitleRoot>/
  <SourceName>-upscaled/
    默認/
      <章节目录>/
        *-upscaled.webp|png|jpeg|avif
      <章节目录>.epub
    单行本/
      <卷目录>/
        *-upscaled.webp|png|jpeg|avif
    <其它分组>/
      <章节目录>/
        *-upscaled.webp|png|jpeg|avif
      <章节目录>.epub

  <SourceName>-output/
    <卷epub>.epub
    <作品名 - 话范围>.epub
    .manga_epub_automation_merge_manifest.json
    .manga_epub_automation_merge_order.json
```

## 快速开始（CLI）

1. 从模板复制依赖配置：

```powershell
Copy-Item .\manga_epub_automation.deps.template.json .\manga_epub_automation.deps.json
```

2. 编辑 `manga_epub_automation.deps.json`，填写本机路径：

- `python_exe`
- `backend_script`
- `models_dir`
- `kcc_exe`

3. 先做计划预检：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -PlanOnly
```

4. DryRun 验证：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -DryRun
```

5. 正式执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>"
```

## 常见阶段组合

```powershell
# 仅超分
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipEpubPackaging -SkipMergedEpub

# 仅章节 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipMergedEpub

# 仅总集合并
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipEpubPackaging

# 额外分组：按话处理并允许各自合并（作品名 - 文件夹名.epub）
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -ExtraGroupHandlingMode chapter_merge
```

## Merge 显式顺序文件

默认路径：
`<TitleRoot>\<SourceName>-output\.manga_epub_automation_merge_order.json`

格式：

```json
{
  "version": 1,
  "chapters": ["117话", "118话", "119话"]
}
```

规则：

- `chapters` 的顺序就是 merge 优先顺序
- 未覆盖章节会按自动排序追加到末尾
- 文件非法（JSON 或字段错误）会阻断执行

## GUI

构建并运行：

```powershell
dotnet restore .\MangaEpubAutomation.Gui.sln
dotnet run --project .\gui\MangaEpubAutomation.Gui\MangaEpubAutomation.Gui.csproj
```

## 运行产物

`logs/` 目录会生成：

- `run_plan_*.json`
- `run_result_*.json`
- `manga_epub_automation_run_*.log`
- `latest_run_plan.json`
- `latest_run_result.json`

GUI 通过 `-GuiMode` 消费 `PIPELINE_EVENT:` 事件流。

## 近期更新

- 支持“仅默认组、无单行本”源目录结构，流程可正常执行。
- 支持自动识别 `默認/` 中的“卷目录”并按单行本方式处理。
- 支持处理额外分组目录（非 `默認/单行本`）：按话打包章节 EPUB，不参与 merge。
- 总集 EPUB 命名优化：优先用章节名中可解析的“数字+话/話”作为范围端点，避免 `最终话` 等特殊命名导致范围偏差。
- 新增 MangaJaNai 灰度检测阈值配置（`WorkflowOverrides.GrayscaleDetectionThreshold`），GUI 可直接设置。
- 灰度阈值范围统一为 `0-24`（GUI 输入、提示文案、保存时范围校验一致）。
- GUI 已禁用“鼠标滚轮调整数值”，避免误触导致参数被改动。
- GUI 在进程非 0 退出时会立即弹窗提示（退出码、`latest_run_result.json`、最新日志路径、stderr 摘要），无需手动翻日志。

## 开源与合规说明

- 本仓库只包含自动化脚本与 GUI 代码。
- 请确保你有权处理对应漫画资源。
- 第三方依赖请遵循其各自许可证。

## License

本仓库代码采用 MIT，见 [LICENSE](LICENSE)。
