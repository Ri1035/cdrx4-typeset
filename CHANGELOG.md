# 变更记录（CHANGELOG）

版本号规则：`主.次.修订`。**修订号**用于修 bug（不改变对外行为），**次版本号**用于新增功能或改变安装方式。

---

## v1.0.0 — 2026-09-28

**首个版本**：把工作区 `[排版]` 组（`GlobalMacros`，**加密**，protection=1）的 6 个按钮
用可读源码重写成独立插件 `TypesetToolkit`，工具栏名「排版工具」。

与同作者的 `cdrx4-toolkit`（「增强工具」，9 按钮）是**两个独立产品**：项目名、工具栏名、
按钮 guid 均不冲突，可同时安装。

### 新增

6 个功能，每个 = 一个模块 = 一个按钮：

| 按钮 | 模块 | 实现要点 |
|---|---|---|
| **拼版** | `src\M_Impose.bas` | `StepAndRepeat` + `CreateShapeRangeFromArray` 矩阵铺开；原对象留在左上角 |
| **裁切线** | `src\M_CropMark.bas` | 包围盒四角各 2 条线段（共 8 条），套版色 0.25mm，末尾群组 |
| **书脊计算** | `src\M_Spine.bas` | 公式 `厚度 = 0.135 × 克重/100 × 页数/2`，支持正算/反算 |
| **中心对齐** | `src\M_Align.bas` | 所有选中对象收敛到**最底层对象**的中心 |
| **按间距分布** | `src\M_Distribute.bas` | 固定间距（火车排列）+ 平均分布两种模式 |
| **删除线段** | `src\M_DelSegment.bas` | 曲线选中节点删除 / 直线段对象整体删除 |

配套：

- **`src\M_Util.bas`** 公共库：单位换算、选择集校验、包围盒、最底层对象、曲线/直线段判定、套版色。
- **`src\M_Install.bas`** 工具栏注册（`InstallToolbar` / `UninstallToolbar` / `DiagToolbar`），
  `CmdList()` 定义 6 按钮的命令、中文标题与提示。
- **`src\M_Test.bas`** 自检（`Probe` / `Diag` / `SelfTest`），写 `%TEMP%\typeset_*.log`。
- **一键安装器** `dist\安装排版工具.vbs`：GMS 以 base64 内嵌，单文件双击即装；
  把工具栏标记写进 `DRAWUIConfig.xml`（改前备份 `.cdrx4bak`）。
- **一键卸载器** `dist\卸载排版工具.vbs`：删 GMS + 清工作区标记。
- **不装启动钩子**（`src\_startup.txt` 首行 `--none`）。

### 验证记录（本机，CorelDRAW X4）

| 项 | 结果 |
|---|---|
| 第 1 层 编译 + 会话内自检 | `build_gms.vbs` → **BUILD OK**（exit 0），`logs\build.log` |
| 书脊正算 克重80/页数100 | **10.8 mm** ✅ |
| 书脊反算 厚度10.8/页数100 | **80 g/m²** ✅ |
| 中心对齐 | moved=**2**，中心收敛 **(50,50)** ✅ |
| 按间距分布 | 实测间距 = 输入 5mm 换算值 **0.1969**（文档单位英寸）✅ |
| 删除线段（对象 / 节点） | **True / 1 / 6→6** 与 **4→2** ✅ |
| 裁切线 | 画出 **8** 条 ✅ |
| 拼版 3×2 | 总数 **6**，对象数净 **+5** ✅ |
| 工具栏自检 | `found=True`，`controls=6`，中文标题与提示逐条比对通过 ✅ |
| 第 3 层 真实启动 | 手动启动 X4：12 线程 / 75.6MB / responding / 新增崩溃事件 **0** → **LAUNCH PASS** |
| 工具栏落屏 | 枚举进程顶层窗口，出现 `title=排版工具` 的工具栏窗口 → **TOOLBAR PASS** |
| 安装器 | 沙箱 + 真实安装双验证；2 份 `DRAWUIConfig.xml` 均 `load=True`、`itemData=6` |
| 产物 | `dist\TypesetToolkit.gms` 138770 字节 |

### 已知限制

- X4 的 `.bas` 无法导入 UserForm，「书脊计算」用 InputBox 顺序问答代替原窗体（流程等价，外观不同）。
- 裁切线只做标准四角角线，不含原 `AutoCutLines` 的旅行商式智能排线。
- 拼版只做当前页内的矩阵复制，不跨页。
- 所有 mm 输入按当前文档单位换算。