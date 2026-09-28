# TypesetToolkit · 排版工具（CorelDRAW X4 插件）

把工作区 `[排版]` 组的 6 个加密按钮，用可读源码重写成一个可安装的 X4 插件。

工具栏名：**排版工具**，6 个按钮：
**拼版 · 裁切线 · 书脊计算 · 中心对齐 · 按间距分布 · 删除线段**

---

## 一、安装

1. **先关掉 CorelDRAW**（安装器检测到 X4 在运行会拒绝执行——正在运行的 X4 退出时会回写工作区，把刚写进去的工具栏冲掉）。
2. 双击 **`dist\安装排版工具.vbs`**。
3. 打开 CorelDRAW，「排版工具」工具栏自动出现，6 个按钮直接可用。

安装器做两件事：

- 把 `TypesetToolkit.gms` 写进 X4 的用户 GMS 目录
  （`%APPDATA%\Corel\CorelDRAW Graphics Suite X4\User Draw\GMS`）；
- 把工具栏定义写进 X4 的**工作区文件** `DRAWUIConfig.xml`（递归处理
  `User Workspace\CorelDRAW\` 下的每一份），这样按钮的中文标题才会显示。
  改写前会留一份 `DRAWUIConfig.xml.cdrx4bak` 备份。

> 如果工作区写入失败（例如目录权限问题），安装器会提示你手动装一次：
> 工具(T) > 宏 > 运行宏 > `TypesetToolkit` > `M_Install` > `InstallToolbar`。

## 二、卸载

双击 **`dist\卸载排版工具.vbs`**：删除 `TypesetToolkit.gms`，并从每份
`DRAWUIConfig.xml` 里移除本插件的工具栏标记。

## 三、用法

除「书脊计算」外，都是**先选中对象再点按钮**。

| 按钮 | 怎么用 |
|---|---|
| **拼版** | 选中对象 → 输入 `列数 行数 列间距(mm) 行间距(mm)`（默认 `3 2 5 5`，分隔符可用空格/`x`/`*`/`,`）→ 按矩阵复制铺开，原对象留在左上角 |
| **裁切线** | 选中对象 → 输入角线长度(mm)、出血间距(mm) → 四角画出 L 形角线（8 条，套版色，自动群组） |
| **书脊计算** | 直接点。模式 1：克重+页数→厚度；模式 2：厚度+页数→克重。公式 `厚度 = 0.135 × 克重/100 × 页数/2` |
| **中心对齐** | 选中多个对象 → 全部收敛到**最底层对象**的中心（底层对象不动） |
| **按间距分布** | 选中多个对象 → 模式 1 固定间距（最左不动，其余依次排开）；模式 2 平均分布（首尾不动，中间等间距） |
| **删除线段** | ① 曲线上选中 ≥2 个节点 → 删掉这些节点及线段；② 选中直线段对象 → 整个删掉 |

每个功能的参数、边界与实测结果见 [docs/功能清单.md](docs/功能清单.md)。

## 四、环境与兼容性

- **仅 CorelDRAW X4**（VBA/GMS 架构，X5+ 的 `.gms` 不通用）。
- 需要 X4 安装时勾选了 VBA。
- 所有长度参数按**当前文档单位**换算。
- **不装启动钩子**：X4 启动期在 `GlobalMacroStorage_Start` 里操作 CommandBars 会崩
  （`CrlFrmWk.dll`，0xc0000005），所以工具栏改为由安装器写工作区文件，装一次长期有效。

## 五、目录结构

```
cdrx4-typeset/
├─ dist/                      交付物
│   ├─ 安装排版工具.vbs          单文件安装器（内含 GMS 的 base64）
│   ├─ 卸载排版工具.vbs          卸载器
│   └─ TypesetToolkit.gms       插件本体（138 KB）
├─ src/                       VBA 源码（9 个模块）
│   ├─ M_Util.bas             公共库（单位换算/选择集/包围盒/颜色）
│   ├─ M_Impose.bas           拼版
│   ├─ M_CropMark.bas         裁切线
│   ├─ M_Spine.bas            书脊计算
│   ├─ M_Align.bas            中心对齐
│   ├─ M_Distribute.bas       按间距分布
│   ├─ M_DelSegment.bas       删除线段
│   ├─ M_Install.bas          工具栏安装/卸载/自检
│   ├─ M_Test.bas             自检（Probe/Diag/SelfTest）
│   └─ _startup.txt           启动钩子占位（首行 `--none` = 不装钩子）
├─ docs/功能清单.md            6 功能明细 + 验证矩阵
├─ docs/优化问题总结.md         交付后遗留的优化点（功能/工程/环境/安全）
├─ tools/                     构建/侦察/验证脚本
│   ├─ real_launch_test.ps1   真实启动不崩检查
│   └─ verify_toolbar.ps1     工具栏是否真的出现在屏幕上
├─ build_gms.vbs              源码 → GMS（需要本机装 X4）
├─ make_installer.ps1          GMS + 文案 → dist/*.vbs
├─ workspace_markup.txt        工具栏标记（6 按钮 + 中文标题）
├─ CHANGELOG.md                变更记录（版本 + 验证结果）
├─ PLAN.md                     工作日志（每个坑一条，含证据）
├─ LICENSE                     MIT
└─ PROGRESS.md                 进度/计划/断点续跑（唯一进度权威）
```

## 六、开发者：重新构建

```powershell
# 1) 改完 src\ 后重新编译（需先关掉 CorelDRAW）
cscript //nologo build_gms.vbs          # 必须看到 ==== BUILD OK ====

# 2) 改了工具栏标记或文案后重新生成安装器
powershell -ExecutionPolicy Bypass -File make_installer.ps1

# 3) 验证
cscript //nologo _test_silent.vbs                              # 免弹窗安装
powershell -ExecutionPolicy Bypass -File tools\real_launch_test.ps1   # 启动健康检查
powershell -ExecutionPolicy Bypass -File tools\verify_toolbar.ps1     # 工具栏是否出现
```

构建日志：`%TEMP%\typeset_selftest.log` / `typeset_diag.log` / `typeset_toolbar.log`。

## 七、复用与署名

- 矩阵复制、裁切线、线段/节点删除的算法借鉴开源项目
  **`hongwenjun/corelvba`**（`Arrange.bas`、`cropline.bas`、
  `SelectLine_to_Cropline.bas`、`Tools.bas`）。
- 工作区内可读源码 **`ZeroBase.gms`**（`ArrangeForm`、`AutoCutLines`）提供了
  `StepAndRepeat` 矩阵铺开与角线生成的参考实现。
- 「书脊计算」的公式直接照搬工作区内可读源码 `Draw\GMS\书脊计算.gms` 的 `UserForm1`。
- 6 个按钮的原宏路径来自工作区泄漏的 `DrawUIConfig.xml`。

本项目为上述来源的**重实现**，非反编译产物。

## 八、版本历史

| 版本 | 日期 | 摘要 |
|---|---|---|
| **v1.0.0** | 2026-09-28 | 首个版本：6 个排版功能 + 工具栏 + 一键安装/卸载器 |

详细变更与验证结果见 [CHANGELOG.md](CHANGELOG.md)；开发过程中的每个坑见 [PLAN.md](PLAN.md)。

**与 `cdrx4-toolkit`（「增强工具」，9 按钮）的关系**：两个独立产品，项目名、工具栏名、
按钮 guid 均不冲突，可同时安装、互不干扰。

许可证：[MIT](LICENSE) · 版权所有 (c) 2026 KOKA (KOKACODA)