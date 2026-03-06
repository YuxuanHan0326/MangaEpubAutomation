# MangaEpubAutomation

一个面向漫画图片目录的自动化流水线：

1. 超分辨率处理（MangaJaNai backend）
2. 单章/单卷 EPUB 打包（KCC）
3. 默认分组章节合并为总集 EPUB（按 `order` 排序）

本仓库同时提供：

- CLI 主入口（PowerShell）
- WPF GUI（Wpf.Ui）

## 适用场景

- 已有结构化漫画图片目录，希望批量超分并打包 EPUB
- 支持增量更新：已存在输出会自动跳过
- 支持在 merge 阶段使用显式顺序文件覆盖默认排序

## 第三方依赖（用户自行安装）

- `copymanga-downloader`（用于获取并组织输入目录）
  - https://github.com/misaka10843/copymanga-downloader
- `MangaJaNaiConverterGui`（提供 python/backend/models 运行环境）
  - https://github.com/the-database/MangaJaNaiConverterGui
- `KCC (comic2ebook)`（用于分章/分卷 EPUB 打包）
  - https://github.com/ciromattia/kcc

说明：

- 本仓库不分发上述依赖本体、模型和漫画内容。
- `kcc` 可执行文件需要用户自行下载并配置路径。

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
└─ logs/                      # 运行时生成（已被 .gitignore 忽略）
```

## 快速开始（CLI）

1. 复制依赖模板：

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

5. 实际执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>"
```

## 阶段控制示例

```powershell
# 仅超分
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipEpubPackaging -SkipMergedEpub

# 仅分章 EPUB
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipMergedEpub

# 仅总集合并
powershell -ExecutionPolicy Bypass -File .\Invoke-MangaEpubAutomation.ps1 -TitleRoot "<TitleRoot>" -SkipUpscale -SkipEpubPackaging
```

## 输入目录约定

`-TitleRoot` 指向单作品根目录，期望结构如下：

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

  <SourceName>-output/
    <单行本epub>.epub
    <作品名 - 话范围>.epub
    .manga_epub_automation_merge_manifest.json
    .manga_epub_automation_merge_order.json
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

- `chapters` 顺序即 merge 优先顺序
- 未覆盖章节会按自动排序追加在末尾
- 文件非法（JSON 或字段错误）会阻断执行

## GUI（Wpf.Ui）

构建/运行：

```powershell
dotnet restore .\MangaEpubAutomation.Gui.sln
dotnet run --project .\gui\MangaEpubAutomation.Gui\MangaEpubAutomation.Gui.csproj
```

页面：

- Run
- Dependencies
- Pipeline Config
- Merge Order
- Logs & Result

## 日志与机器接口

运行会在 `logs/` 生成：

- `run_plan_*.json`
- `run_result_*.json`
- `manga_epub_automation_run_*.log`
- `latest_run_plan.json`
- `latest_run_result.json`

GUI 通过 `-GuiMode` 消费 `PIPELINE_EVENT:` 事件流。

## License

本仓库代码采用 MIT，见 [LICENSE](LICENSE)。
