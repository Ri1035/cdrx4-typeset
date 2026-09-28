# 工作日志（PLAN）

> 每个坑记一条，含**证据**（命令 / 日志 / 报错原文），供以后踩到时直接对号入座。
> 阶段划分与 `cdrx4-toolkit` 一致：侦察 → 功能清单 → 写代码 → 三层验证 → 安装器 → 版本管理。

---

## 1. 目标与来源

复刻工作区 `[排版]` 组（`GlobalMacros`）的 6 个**加密**按钮。

- 参考插件：`C:\Program Files (x86)\CorelDRAW X4\Draw\GlobalMacros.gms`，**protection=1（加密）**，源码不可读。
- 按钮 → 宏路径从工作区泄漏的 `DrawUIConfig.xml` 拿到（6 条）。
- 例外：`书脊计算` 的实现落在**可读的** `Draw\GMS\书脊计算.gms`（`ThisMacroStorage` + `UserForm1`），公式直接可读 → 走**移植**。

## 2. 复用来源（省时省钱的关键）

| 功能 | 复用来源 |
|---|---|
| 拼版 | `ZeroBase.gms` 的 `ArrangeForm`（`StepAndRepeat` 矩阵铺开）、开源 `hongwenjun/corelvba` 的 `Arrange.bas` |
| 裁切线 | `ZeroBase.gms` 的 `AutoCutLines`、开源 `cropline.bas` |
| 删除线段 | 开源 `SelectLine_to_Cropline.bas`、`LinesTool` 的节点操作 |
| 按间距分布 | 开源 `Tools.bas`（对齐/分布/拆分线段） |
| 书脊计算 | `Draw\GMS\书脊计算.gms` 的 `UserForm1`（公式照搬） |

按约定：**不整段照搬**，改写算法并在 README 署名致谢。

## 3. 坑（含证据）

### 3.1 X4 的 VBE 不按文件名给模块命名 —— 缺 `Attribute VB_Name` 就废
**现象**：首次构建 `!! module M_Install is present 0 time(s)` / `!! module M_Test is present 0 time(s)`
→ `module names collided - macro paths would break`。
**原因**：9 个模块里 7 个首行有 `Attribute VB_Name = "M_Xxx"`，`M_Install.bas` / `M_Test.bas` 没有。
X4 的 VBE 对缺 Attribute 行的模块**不**用文件名命名，会取一个自造名 → 组件名校验失败。
**修复**：给这两个文件补 `Attribute VB_Name`（`src\M_Install.bas:1`、`src\M_Test.bas:1`）。
**顺带加固**：`build_gms.vbs` 在名字不符时**转储真实组件名**，否则报错只有「present 0 time(s)」，无从下手。

### 3.2 X4 退出时会自动回写 GMS —— 失败的构建会留下半个工程
**现象**：构建在 `[4]` 步失败后，`%APPDATA%\...\User Draw\GMS\TypesetToolkit.gms` **仍然存在**，
138770 字节（种子 `Emboss.gms` 只有 54290），`LastWriteTime` 比 `CreationTime` 晚 6 秒。
**原因**：`Fail()` 里先 `QuitApp`，X4 关闭时把**已修改的工程**自动保存回目标路径，
而 `DeleteFile` 跑在保存完成之前 → 文件复活。这个半成品下次启动会被 X4 加载并弹编译错误。
**修复**：`Fail()` 改成**循环重删**（最多 20 次 × 500ms，直到文件不再出现），见 `build_gms.vbs` 的 `Fail`。

### 3.3 脚本里不要用带命名参数的 `SetOutlineProperties`
**决定**：裁切线描边改用属性赋值 `sr.Outline.Width = ...` / `sr.Outline.Color.CopyAssign col`。
**原因**：`SetOutlineProperties(..., Color:=col)` 依赖 TLB 命名参数解析，脆弱；
属性写法在同族模块 `M_Seal` 里已验证可用。

### 3.4 拼版 `rows=1` 时结果集漏掉横向副本
**现象**：`all` 初始被赋成 `base`，只有 `rows > 1` 分支才重新赋值 → 只铺一行时最后只选中原对象。
**修复**：横向复制后立刻 `Set all = rowRange`，见 `src\M_Impose.bas`。

### 3.5 启动钩子是雷，默认必须是「无」
**决定**：`src\_startup.txt` 首行 `--none`。
**原因**：`GlobalMacroStorage_Start` 触发时命令栏框架尚未初始化完，在里面建/删 `CommandBars`
会让 X4 直接崩（`CrlFrmWk.dll`，`0xc0000005`），且 `On Error Resume Next` 挡不住访问冲突
（那不是可捕获的 VBA 错误），并留下杀不掉的 0 线程僵尸进程。
**替代方案**：工具栏由**安装器写工作区文件**，装一次长期有效。

### 3.6 构建脚本结尾要主动退出 X4，否则留孤儿进程
**修复**：成功路径末尾也调用 `QuitApp`。本机历史上那批 0 线程僵尸就是这么来的。

### 3.7 静默测试版多打印的 `64` 不是 bug
**现象**：`_test_silent.vbs` 输出末尾多出 ` 64 排版工具`。
**原因**：`MsgBox Fmt(text), vbInformation, TB` 被替换成 `WScript.Echo Fmt(text), vbInformation, TB`，
Echo 会把所有参数打出来，而 `vbInformation = 64`。真实安装器走 MsgBox，不受影响。

### 3.8 本机 5 个 0 线程僵尸 CorelDRW.exe
**证据**：`Get-CimInstance Win32_Process -Filter "Name='CorelDRW.exe'"` →
PID 15148 / 9012 / 3020 / 8316 / 3540，`ThreadCount=0`，`SessionId=1`；`Stop-Process -Force` 无效。
**影响**：不在 ROT，**不阻塞** `CreateObject` 与构建/安装（已实测跑通）；但需**重启系统**才能清掉。

## 4. 三层验证（每次改代码都要重跑）

```powershell
cscript //nologo build_gms.vbs                                            # 第1层：编译 + 会话内自检
powershell -ExecutionPolicy Bypass -File tools\real_launch_test.ps1 -WaitSec 25   # 第3层：真实启动不崩
powershell -ExecutionPolicy Bypass -File tools\verify_toolbar.ps1 -WaitSec 24     # 工具栏是否落屏
```

第 2 层冒烟由 `build_gms.vbs` 内的 `M_Test.SelfTest` 承担（8 个用例全绿）。
`tools\verify_toolbar.ps1` 是本项目**新增**的验证脚本：枚举进程顶层窗口找标题，
把「工具栏到底有没有出现」从人工目测变成可自动判定。

## 5. 版本管理与推送机制

- 维护 `CHANGELOG.md`（版本 + 日期 + 改了什么 + 验证结果）、`README.md`、`PLAN.md`（本文件）。
- **一个功能一个提交**。
- 工作副本**不是** git 仓库；推送靠同级目录的暂存克隆 `_push_stage\`：
  把要发布的文件复制进去 → `git add/commit` → **经用户确认** → `push`。
- **侦察产物一律不进仓库**：`_ref\`、`_recon_src\`、`recon_ws.txt`、`*.log`、`_*.png`、`_push_stage\`。
- 中文提交信息走 `git commit -F <UTF-8 文件>`，避免 PowerShell 的 Latin-1 把中文/文件名搞坏。
- 凭据：OAuth 连接器身份 `KOKACODA` 对 `Ri1035/*` 仓库 **push 权限不足**，
  必须用 `Ri1035` 账号的 token；**token 不写入 `.git/config`**，用完吊销。