# 光谱失配计算

这是一个本地运行的光谱 SPD 自动计算器，用于读取参考光谱、测试光谱和光谱响应 SR，计算 SPD 光谱偏差率、SPC 覆盖率、IEC 分段匹配等级，以及 ISC/CV 修正结果。

## 快速启动

双击 `启动光谱SPD计算器.bat` 即可启动。也可以在当前文件夹运行：

```powershell
.\启动光谱SPD计算器.bat
```

启动失败时，程序会在当前文件夹写入 `SPD_startup_log.txt`，其中包含失败原因。

## 关键文件

- `spectral_spd_gui.py`：主程序，包含读取、计算、界面和导出逻辑。
- `启动光谱SPD计算器.bat`：给日常用户使用的启动入口。
- `start_spd_calculator.ps1`：启动脚本，优先使用 Codex 内置 Python，找不到时回退到系统 Python。
- `使用说明_SPD计算器.md`：面向使用者的操作说明和公式说明。
- `video_spd_intro/`：用 HyperFrames 制作的 SPD 计算器演示视频项目，包含素材、预览图和 MP4 渲染结果。
- `AM1.5标准光谱数据280-4000nm.xlsx`：默认参考光谱文件，程序会自动识别 `AM1.5G` / `AM1.5D` 列；手动选择其他参考表格时，会根据表头提供可选参考光谱列，例如 `AM0`。
- `AM0.xlsx`：可手动选择的 AM0 参考光谱文件。
- `1027SR.xlsx`：默认光谱响应 SR 文件；程序也兼容 `1025SR.xlsx`、`1025SR*.xlsx` 和文件名包含 `光谱响应` 的 SR 表格。

## 输入与输出

程序支持读取 `.xlsx`、`.xlsm`、`.xls`、`.csv`、`.txt`、`.tsv`、`.asd` 表格。文本表格支持逗号、Tab、分号或空格分隔，并会尝试 `utf-8-sig`、`utf-8`、`gbk`、`gb18030` 编码。`添加 IV` 可以一次选择多个 IV 表格；文本 IV 文件会识别电压/电流表头，纯数字两列文本也会按电压 V、电流 A 读取并换算为 mA。

导出结果可以保存为 `.xlsx` 或 `.csv`。导出 `.xlsx` 时会包含 `Summary`、`Bands`、`Data`、`Charts` 等工作表；如果已经完成 ISC 修正，还会包含 `CV_Groups`、`ISC_Correction`、`SR_Data`。导出结果时，程序会同步保存同名 `.inp` 输入文件，便于下次恢复现场。

单独导出 `光谱相关数据.xlsx` 时，`Summary` 页会先显示一张与界面一致的所有光谱对比图，下方汇总每组光谱的 SPD、SPC、MG、MMF；每组光谱还会生成独立页面，包含单组对比图和原始光谱数据。

多组光谱同时绘图时，程序会使用更丰富的颜色池，并在颜色循环后自动叠加虚线、点线等线型；右侧会预留图例区域，图例数量增多时会自动缩小图例字号并在需要时分列显示。`光谱记录` 会显示每组测试光谱对应的曲线样式，并提供 `独立显示` 复选框；可通过 `批量显示` 或 `独立显示` 表头批量全选/全不选。选中光谱记录后，空格切换显示、`F2` 改名、`Enter` 高亮；双击图例中的测试光谱也可多选高亮，被选曲线保持醒目，其余测试曲线变淡，参考曲线保持不变。

## 自检

核心计算自检不打开界面：

```powershell
$env:SPD_SELF_TEST = "1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\start_spd_calculator.ps1
$env:SPD_SELF_TEST = $null
```

GUI 冒烟检查会打开界面并自动关闭：

```powershell
$env:SPD_GUI_SMOKE = "1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\start_spd_calculator.ps1
$env:SPD_GUI_SMOKE = $null
```
