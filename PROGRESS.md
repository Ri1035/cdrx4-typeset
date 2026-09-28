# 排版工具 TypesetToolkit —— 全量计划 · 进度 · 断点续跑

> 本文件是**唯一的进度权威**。上下文被压缩后，先读本文件再动手，不要重新侦察、
> 不要重跑已标 ✅ 的步骤。每完成一步就把状态改成 ✅ 并补上验证证据。
>
> 最后更新：2026-09-28 17:0x

---

## 0. 目标与交付物

复刻工作区 `[排版]` 组（GlobalMacros）的 **6 个加密按钮**，用可读源码重写，
产出可安装的 CorelDRAW X4 插件。

| 按钮 | 宏入口 | 语义（已侦察确认） |
|---|---|---|
| 拼版 | `M_Impose.Impose` | 选中对象按 行×列 矩阵复制铺开（可选行/列间距） |
| 裁切线 | `M_CropMark.CropMark` | 围绕选中对象包围盒画四角 L 形角线 |
| 书脊计算 | `M_Spine.SpineCalc` | 按克重×页数算书脊厚度，或反算克重 |
| 中心对齐 | `M_Align.CenterAlign` | 所有选中对象收敛到**最底层对象**的中心 |
| 按间距分布 | `M_Distribute.DistributeBySpacing` | 按固定间距或平均间距排成一行 |
| 删除线段 | `M_DelSegment.DelSegment` | 删除直线段对象，或曲线上选中节点间的线段 |

**最终交付**：`dist\安装排版工具.vbs`（单文件，双击即装）+ `dist\卸载排版工具.vbs` + 本目录文档。

---

## 1. 环境事实（已核实，勿再探）

| 项 | 值 |
|---|---|
| X4 程序 | `C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDRW.exe` |
| X4 根 | `C:\Program Files (x86)\CorelDRAW X4`（子目录含 `Draw\GMS`、`Programs`） |
| 类型库 | `C:\Program Files (x86)\CorelDRAW X4\Programs\CorelDraw.tlb` |
| 种子 GMS | `C:\Program Files (x86)\CorelDRAW X4\Draw\GMS\Emboss.gms`（54290 B，未加密） |
| 安装目标 | `%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS\TypesetToolkit.gms` |
| 工作区配置 | `%APPDATA%\Corel\...\User Workspace\CorelDRAW\` 下**递归找** `DRAWUIConfig.xml`，当前 2 个：<br>· `Adobe(R) Illustrator(R)\DrawUIConfig.xml`（X4 当前生效的工作区）<br>· `_default\DRAWUIConfig.xml` |
| 项目名 | `TypesetToolkit`（必须与 `M_Install.PRJ_NAME`、`dynamicCommand` 前缀三者一致，否则宏路径全废） |
| 工具栏名 | `排版工具`，guid `cf4b97c9-1db9-4c5f-afea-46fa2d54b273` |

**本机进程坑（重要）**：
- 有 **5 个 0 线程僵尸 `CorelDRW.exe`（PID 15148/9012/3020/8316/3540，Session 1）**，
  `Stop-Process` 无效，**必须重启系统才能清掉**。僵尸不在 ROT，不阻塞 `CreateObject`，
  所以**构建/安装仍可跑通**（已验证），但会残留。
- 用户已明确：**进程干扰时由用户手动重启电脑，助手禁止重启电脑**。

---

## 2. 已完成（✅ = 已验证，带证据）

### ✅ 2.1 源码（`src\`，9 个模块 + 1 个占位）

| 文件 | 说明 |
|---|---|
| `M_Util.bas` | 公共库：单位换算 `MmToDoc/DocToMm`、`RequireSelection`、`SelectionBounds`、`BottomMostOf`、`ShapeCurve`、`IsLineSegment`、`MakeCMYK/MakeRGB`、`RegistrationColor` |
| `M_Spine.bas` | 书脊：`SpineCalc` / `SpineThicknessMm` / `WeightFromSpine`（公式 `0.135 × 克重/100 × 页数/2`，来自可读源码） |
| `M_Align.bas` | 中心对齐：`CenterAlign` / `AlignCentersCore`（基准 = 最底层对象） |
| `M_Distribute.bas` | 按间距分布：`DistributeBySpacing` / `DistributeBySpacingCore`（固定间距 + 平均分布两模式） |
| `M_DelSegment.bas` | 删除线段：`DelSegment` / `DelSegmentCore`（`NodeRange.Delete` + 直线段对象 `Delete`） |
| `M_CropMark.bas` | 裁切线：`CropMark` / `CropMarkCore`（4 角 × 2 线 = 8 条，套版色 0.25mm，最后群组） |
| `M_Impose.bas` | 拼版：`Impose` / `ImposeCore`（`StepAndRepeat` + `CreateShapeRangeFromArray`） |
| `M_Install.bas` | 工具栏安装/卸载/自检：`InstallToolbar` / `UninstallToolbar` / `DiagToolbar`，`CmdList` 定义 6 按钮 |
| `M_Test.bas` | 自检：`Probe` / `Diag` / `SelfTest`，写 `%TEMP%\typeset_*.log` |
| `_startup.txt` | 首行 `--none`（**不装启动钩子**，避免 X4 启动期建 CommandBars 崩溃） |

**关键约定**：每个 `.bas` 首行必须是 `Attribute VB_Name = "M_Xxx"`。
X4 的 VBE 对**缺** Attribute 行的模块不按文件名命名，会取一个自造名 →
组件名校验失败、宏路径全废。`M_Install.bas` / `M_Test.bas` 曾因此构建失败，已补上。

### ✅ 2.2 构建链

| 文件 | 说明 |
|---|---|
| `build_gms.vbs` | 已改为 TypesetToolkit：模块表、`p.Name="TypesetToolkit"`、日志名 `typeset_*.log`、`RunMacro("TypesetToolkit", ...)`。新增：组件名不符时**转储真实组件名**；`Fail()` 循环重删（X4 退出时会自动回写半个 GMS）；成功路径末尾 `QuitApp`（不再留孤儿进程） |
| `workspace_markup.txt` | 6 个 `<itemData>`（含新 guid + `dynamicCommand="TypesetToolkit.M_Xxx.Yyy"` + 中文 caption/tooltip）+ `排版工具` 工具栏 + 可见性 |
| `installer_template.txt` | `GMS_NAME=TypesetToolkit.gms` / `PRJ_NAME=TypesetToolkit`；Strip 正则改 `TypesetToolkit\.`；输出名 `安装排版工具.vbs` |
| `uninstaller_template.txt` | 同上改名；Strip 正则 `[^"]*TypesetToolkit[^"]*`；输出名 `卸载排版工具.vbs` |
| `installer_msgs.txt` / `uninstaller_msgs.txt` | 文案改名（排版工具 / 6 按钮 / TypesetToolkit 宏路径） |
| `make_installer.ps1` | 路径与产物名改为 TypesetToolkit；生成 `dist\安装排版工具.vbs`、`dist\卸载排版工具.vbs`，并落一份 `_test_silent.vbs`（MsgBox→Echo，供无人值守测试） |

### ✅ 2.3 三层验证 —— **全部通过**

**BUILD**：`cscript //nologo build_gms.vbs` → **exit 0 / BUILD OK**
（日志：`logs\build.log`）

**smoke（M_Test.SelfTest，全部 `err=0`）**：

| 项 | 期望 | 实测 |
|---|---|---|
| 书脊 页数100/克重80 | 10.8 mm | 10.8 ✅ |
| 书脊 反算 厚度10.8/页数100 | 80 g | 80 ✅ |
| 中心对齐 moved / c | 2 / (50,50) | 2 / (50,50) ✅ |
| 按间距分布 gap | = MmToDoc(5) | 0.1969（文档单位为英寸）✅ |
| 删除线段（对象） | isline=True, n=1 | True, 1 ✅ |
| 删除线段（节点） | 4→2 | 4→2 ✅ |
| 裁切线 | n=8 | 8 ✅ |
| 拼版 3×2 | total=6, shapes +5 | 6, +5 ✅ |

**toolbar（M_Install.DiagToolbar，`err=0`）**：`found=True, controls=6`，
6 个按钮的中文 caption 与 tooltip 逐条比对通过。

### ✅ 2.4 产物

- `dist\TypesetToolkit.gms` — 138770 B（种子 54290 B → 编译后增长，`sizeAfter > sizeBefore` 已校验）
- `dist\安装排版工具.vbs` — 211406 B
- `dist\卸载排版工具.vbs` — 17019 B
- 安装目标处已有 `TypesetToolkit.gms` 138770 B

---

## 3. 待办（按顺序，每条都有验收标准）

### ✅ 3.1 安装器沙箱验证
`%TEMP%\ts_ws_test\` 造真实工作区副本，`_test_silent.vbs` 里
`gSh.ExpandEnvironmentStrings("%APPDATA%")` → `...("%TEMP%") & "\ts_ws_test"`、
`If CorelRunning() Then` → `If False Then`，跑通。
**结果**：`load=True`，`itemData=6`，工具栏 guid ×2，`commandBarData` 结构正确。
（静默版尾部多出的 ` 64 排版工具` 只是 `MsgBox → Echo` 把 `vbInformation`(=64)
和标题一起打出来了，非缺陷。）

### ✅ 3.2 真实安装
**结果**：GMS 目录 + **2 份** `DRAWUIConfig.xml` 全部写入成功，两份 `load=True`、
`itemData=6`、工具栏 guid ×2，并生成 `.cdrx4bak` 备份。

### ✅ 3.3 真实启动验证
- `tools\real_launch_test.ps1 -WaitSec 25` → **LAUNCH PASS**
  （pid 20768：12 线程 / 75.6MB / responding=True / 新增崩溃事件 = 0）
- 优雅关闭后 X4 **重写了当前生效工作区**（`_default\DRAWUIConfig.xml`，mtime 17:00:37）
  且**完整保留**我们的标记 → X4 已接受该工具栏定义
- 新增 `tools\verify_toolbar.ps1`（枚举进程顶层窗口找标题）→ **TOOLBAR PASS**
  （出现 `vis | class=Afx:* | title=排版工具` 的工具栏窗口）
- 关闭后无新增僵尸

### ✅ 3.4 文档定稿
- `README.md`：安装/卸载/6 功能用法/兼容性/目录/构建/署名
- `docs\功能清单.md`：6 功能逐条（入口、参数、输入容错、行为、边界）+ 验证矩阵 + 已知限制
- **验收**：文案取自源码实际提示语，数值取自实测日志，无臆测

### ✅ 3.5 收尾
已清理：`%TEMP%\ts_ws_test`、`_probe_ref*.log`、`_recon_vbe.log`、`tools\_inst_test.vbs`。
已修正 `.gitignore` 里的旧名（`dist/CDRX4Toolkit.gms` → `dist/TypesetToolkit.gms`）。
保留：`logs\build.log`（BUILD OK 证据）、`_launch.log`、`_toolbar_check.log`（LAUNCH/TOOLBAR 证据）、
`_recon_src\`、`_ref\`、`recon_ws.txt`（署名与复用来源）。

---

## 3.6 交付确认（全部完成）

| 交付物 | 路径 | 状态 |
|---|---|---|
| 插件本体 | `dist\TypesetToolkit.gms`（138770 B） | ✅ 已编译、已装到 X4 |
| 安装器 | `dist\安装排版工具.vbs` | ✅ 沙箱 + 真实安装双验证 |
| 卸载器 | `dist\卸载排版工具.vbs` | ✅ |
| 说明 | `README.md` | ✅ |
| 功能清单 | `docs\功能清单.md` | ✅ |
| 进度权威 | `PROGRESS.md`（本文件） | ✅ |
| 验证脚本 | `tools\real_launch_test.ps1`、`tools\verify_toolbar.ps1` | ✅ 可重复运行 |

**三层验收结论**：BUILD OK → smoke 全绿（8 项）→ LAUNCH PASS + TOOLBAR PASS。
插件已安装在本机，**打开 CorelDRAW 即可看到「排版工具」工具栏**。

> 唯一遗留：本机 5 个 0 线程僵尸 `CorelDRW.exe` 需**重启系统**清理（不影响插件使用）。

---

## 4. 关键决策 / 已踩的坑（别重复踩）

1. **不加启动钩子**。X4 启动期在 `GlobalMacroStorage_Start` 里碰 CommandBars 会崩
   （`CrlFrmWk.dll`，0xc0000005），且 `On Error Resume Next` 挡不住访问冲突。
   工具栏改由**安装器写工作区 XML**（`DRAWUIConfig.xml`），装一次长期有效。
2. **VBE 只吃 ANSI**。导入前用 `ADODB.Stream` 把 UTF-8 源码转 GBK 临时副本。
3. **模块必须有 `Attribute VB_Name`**（见 2.1）。
4. **`SetOutlineProperties(..., Color:=)` 不要用**，改属性赋值
   `sr.Outline.Width = ...` / `sr.Outline.Color.CopyAssign col`（`M_Seal` 已验证）。
5. **X4 退出会自动回写 GMS**，失败的构建会把半个工程丢进自动加载目录 →
   下次启动编译报错。`Fail()` 必须循环重删（已实现）。
6. **拼版 `rows=1` 分支**：`all` 必须指向整行，否则只选中原对象（已修）。
7. **裁切线会群组**：8 条线 → 1 个群组，对象数净 **+1**（自检文案已按此写）。
8. **文档单位随文档**：自检临时文档单位是**英寸**，故 5mm = 0.1969。
9. **`RunMacro` 用 `app.GMSManager.RunMacro("TypesetToolkit", "M_Xxx.Yyy")`**，
   `CreateObject("CorelDRAW.Application.14")` 在 X4 上有效。

---

## 5. 断点续跑（上下文被压缩后照做）

```
1. 读本文件，看 §3 里第一个 ⬜
2. 环境自检：
   Get-CimInstance Win32_Process -Filter "Name='CorelDRW.exe'" |
     Select-Object ProcessId, ThreadCount, SessionId
   —— 只允许「0 线程僵尸」存在；有非僵尸实例就先让用户关掉
3. 构建（改了 src 才需要）：
   cscript //nologo build_gms.vbs   → 必须 BUILD OK
   产物日志：%TEMP%\typeset_selftest.log / typeset_diag.log / typeset_toolbar.log
4. 安装：cscript //nologo dist\安装排版工具.vbs
5. 重新生成安装器（改了 markup/msg 才需要）：
   powershell -ExecutionPolicy Bypass -File make_installer.ps1
```

**不要**重做：侦察（`recon_ws.txt` / `_recon_src\`）、类型库解析（`tools\tlb_dump.ps1`）、
已通过的构建与自检。**要改**代码时只改 `src\`，然后走第 3 步。