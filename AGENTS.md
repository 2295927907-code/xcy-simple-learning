# AGENTS.md

## 项目定位

本项目是单文件 Python 桌面工具，核心入口是 `spectral_spd_gui.py`。它读取参考光谱、测试光谱、SR 和 IV/ISC 数据，输出 SPD、SPC、IEC 分段等级以及 ISC/CV 修正结果。

## 本地约定

- 所有中文 Markdown 和 Python 文件按 UTF-8 处理；在 PowerShell 中读取中文文件时显式使用 `-Encoding UTF8`。
- `spectral_spd_gui.py` 同时包含计算、GUI 和导出逻辑。修改导出字段、工作表名、默认文件名、单位或公式时，同步更新 `README.md` 和 `使用说明_SPD计算器.md`。
- `video_spd_intro/` 是独立的 HyperFrames 演示视频项目；编辑视频构图时遵守该目录内的 `AGENTS.md`、`CLAUDE.md` 和 `DESIGN.md`，修改后运行 `npm run check`。
- 默认参考光谱通过 `*AM1.5*.xlsx` 查找；手动选择其他参考表格时，程序会从表头识别可用参考列，例如 `AM0`。默认 SR 依次匹配 `1027SR.xlsx`、`1027SR*.xlsx`、`1025SR.xlsx`、`1025SR*.xlsx`、`*光谱响应*.xlsx`。
- `启动光谱SPD计算器.bat` 是面向用户的启动入口；`start_spd_calculator.ps1` 是可脚本化入口。
- 源数据文件、`.inp` 输入状态文件、PDF 标准文件、视频素材和用户导出的结果都视为用户数据。不要擅自删除、重命名或覆盖。

## 验证命令

核心计算自检：

```powershell
$env:SPD_SELF_TEST = "1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\start_spd_calculator.ps1
$env:SPD_SELF_TEST = $null
```

GUI 冒烟检查：

```powershell
$env:SPD_GUI_SMOKE = "1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\start_spd_calculator.ps1
$env:SPD_GUI_SMOKE = $null
```
