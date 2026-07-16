# Visio-F

Visio-F 是一个用于 Microsoft Visio 自动化绘图的仓库，当前版本包含 Codex 全局 Visio 自动化能力的代码副本，以及一份横版研发测试流程泳道图的最终交付文件。

## 仓库内容

- `.codex/visio-automation-skill/`：Visio 自动化脚本与技能说明。
- `rd-test-horizontal-swimlane-arrows-on-port.json`：横版研发测试流程图结构化绘图规格。
- `rd-test-horizontal-swimlane-arrows-on-port.vsdx`：最终 Visio 源文件。
- `rd-test-horizontal-swimlane-arrows-on-port.png`：最终导出的预览图。

## 使用方式

在 Windows PowerShell 中执行：

```powershell
.\.codex\visio-automation-skill\scripts\new_visio_swimlane_diagram.cmd `
  -SpecPath .\rd-test-horizontal-swimlane-arrows-on-port.json `
  -OutputPath .\rd-test-horizontal-swimlane-arrows-on-port.vsdx `
  -Visible `
  -KeepOpen `
  -UpdateExisting `
  -ExportPng
```

参数说明：

- `-Visible`：显示 Visio 操作窗口。
- `-KeepOpen`：绘制完成后保留 Visio 打开。
- `-UpdateExisting`：优先复用当前 Visio 画布，只清理并重绘自动化生成的图形。
- `-ExportPng`：同步导出 PNG 预览图。

## 绘图规则

自动化脚本内置以下默认规则：

- 保持横版画布和泳道布局。
- 模块之间预留走线距离，避免图形重叠。
- 连接线使用正交线，从模块边缘中心或角点引出。
- 连接线不穿入模块内部，不复用同一段线作为多条流程线。
- 箭头贴近目标模块边缘，线段终点与箭头方向保持一致。
- 字体根据模块尺寸和规则说明区域自适应放大。

## 版本库说明

Visio 临时锁文件、系统缓存和日志不会提交到 GitHub。最终交付文件应保留 `.json`、`.vsdx` 和 `.png` 三类文件。
