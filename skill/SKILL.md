---
name: vb6ide-api
description: 操作 VB6 IDE 的 REST API 工具集。当需要读写 VB6 工程代码、增删模块、过程级修改、查看工程/组件信息、或自动化操作正在运行的 VB6 IDE 时使用。所有请求必须通过 skill.js 命令发起，禁止自行构造 curl/fetch 请求。
---

# vb6ide-api 操作技能

VB6 IDE 插件（InfoAddin）在 IDE 进程内提供 REST API。本技能通过 `..js` 固化全部请求方式——**一律用 skill.js 命令，不要自己构造 curl / fetch / 原始 HTTP 请求**（Windows 下 curl 单引号、编码、转义均易出错，请求质量不稳定）。

## 使用方式

```bash
cd <项目根目录>
node .claude/skills/vb6ide-api/skill.js <command> [args]
```

### 无 nodejs 环境（VB6 开发机常见）→ 用 cscript 版

```cmd
cscript //nologo .claude\skills\vb6ide-api\skill.vbs <command> [args]
```

`skill.vbs` 与 skill.js **命令完全同构、请求规范一致**（端口探测、非 ASCII 转 \uXXXX、CRLF 归一、代码体只从文件读、声明头拦截、信封判定），仅依赖 Windows 自带的 MSXML2.ServerXMLHTTP 与 FSO，零第三方依赖、零 PowerShell 依赖。判定优先级：有 node → skill.js；无 node → skill.vbs；两者都按同一套命令调用，输出格式一致。
vbs 版差异：① stdin（`-`）不可用，代码体/JSON 一律写临时文件后传路径（但内联 JSON 字面量如 `{"Caption":"X"}` 可直接作参数，等价 skill.js 的 readJsonFrom 内联支持）；② `ctrl-set` 现已同 skill.js 自动包 `{"props":...}`，props 文件写内层对象（如 `{"Caption":"X","Left":100}`）即可，不再需手写 `{"props":{...}}` 包装；③ `form-style`/`ctrl-set` 第三参数均支持内联 JSON（同 js）。

公共项：`--full` 全量输出（大响应默认截断 60000 字符）、`--max N` 调整截断、环境变量 `VB6IDE_PORT` 指定端口（缺省自动探测 8306 起顺延）。出错时 skill.js 以非零退出码退出并给出 `HTTP 状态 + message`，先读错误再决定下一步。

## 命令清单

| 命令 | 用途 |
|---|---|
| `status` | 探测端口与服务状态（**所有任务的第一步**；服务不可达时停止并告知用户开 IDE/启动服务） |
| `project` / `components` / `snapshot` / `ide` | 工程信息只读；`snapshot` 全量快照用于全局诊断（输出大，确认需要时才用） |
| `code-get <mod> [--start N --end M]` | 读模块代码，行号 1 基准含声明区 |
| `module-create <name> [type]` | 新建模块；type：module/class/form/mdiform/usercontrol/proppage，缺省 module |
| `code-put <mod> <file\|-> [--type t]` | 整体写入模块（不存在则创建）；代码体**从文件或 stdin 读取**，禁止 shell 内联 |
| `module-delete <mod>` | 删除整个模块（不可逆） |
| `module-import <绝对路径>` | 从磁盘导入（.bas/.cls；设计器文件不支持，服务端明确报错） |
| `procs <mod>` | 列出模块内全部过程（名称/类型/propertyKind/作用域/行号） |
| `proc-get <mod> <proc>` | 读单个过程；Property 加 `.Get`/`.Let`/`.Set` 后缀 |
| `proc-put <mod> <proc> <file\|->` | 整体替换过程代码 |
| `proc-delete <mod> <proc> [--cleanup]` | 删除过程（不可逆），`--cleanup` 清理遗留空行 |
| `lines-get <mod> [--start --end]` | 读行范围（语义同 code-get） |
| `lines-insert <mod> --start N <file\|->` | 第 N 行前插入多行 |
| `lines-replace <mod> --start N --end M <file\|->` | 替换行范围（空代码=删除该范围） |
| `lines-delete <mod> --start N --end M` | 删除行范围（不可逆；start/end 必须显式给） |
| `refs` / `ref-add` / `ref-del <guid>` / `typelibs [kw]` | 引用库管理；ref-add 三模式：`--guid {G} [--major --minor]` / `--path P` / `--name 名` |
| `form-get <form>` | 窗体信息+控件清单（**坐标单位 twips**，15 twips=1 像素） |
| `ctrl-add <form> <name> <cls> [--left --top --width --height --caption]` | 添加控件 |
| `ctrl-set <form> <ctrl> <props.json\|->` | 设控件属性；文件写内层对象 `{"Caption":"X","Left":100}`（自动包 {props:...}，支持内联 JSON） |
| `ctrl-del <form> <ctrl>` | 删除控件（不可逆） |
| `form-align <form> --mode M --controls a,b,c [--value N]` | 对齐/等距；mode：left/right/top/bottom/center-h/center-v/width/height/space-h/space-v |
| `form-style <form> <style.json\|->` | 窗体样式；文件形如 `{"caption":"X","width":6000,"startUpPosition":1}` |
| `win-show [immediate\|locals\|watch\|project\|properties]` | 打开 IDE 工具窗口 |
| `save` | 保存当前工程（工程组只存 Active） |
| `debug-status` / `debug-run` / `debug-stop` | 调试三态；run 语义按状态分（design=启动/break=继续；run 态返回 409） |
| `bps` / `bp-set <mod> <line>` / `bp-del <mod> <line>` / `bp-clear` | 断点（**服务端自记账**，IDE 手工断点不可见；**bp-set 是 toggle**，同行重复调用=取消） |
| `debug-output` | 读立即窗口输出 |
| `compile` | **影子编译**：隐含 save 后按磁盘快照在第二 VB6 实例 `/make`（约 3-10 秒，命令内部 1s 轮询至完成）；成功一句话+耗时，失败逐条 `模块 行号 描述` |
| `guide` | **拉服务端下发的 VB6 编程规范**（会话内首次写码前必读，见下节）：VB6≠VBA/VBS 规则、VBA 常量混入黑名单 |

## 写码前置：编程规范注入（LLM 自判，硬性）

- **判断规则**：若本会话上下文中尚未出现过 VB6 编程规范内容（`guide` 输出或响应 `recommendations.rule`），则**任何写码操作（code-put / proc-put / lines-insert / lines-replace / ctrl-add 等）之前必须先跑一次 `guide`**。
- 规范来自服务端 GET /api/guidelines（0.1.17 起，随 DLL 版本演进，无需更新 skill 即可获得最新版）：目标语言定位、规则清单、VBA 常量混入黑名单（如 `vbFromDatabase` 是 Access VBA 专属，VB6 编译报「变量未定义」，改用 `vbFromUnicode`）。
- 所有 API 响应信封顶层带 `recommendations` 字段：`lang`（语言定位）、`rule`（一句核心规范）、`guidelines`（完整规范端点指引）、`skillLatest`（服务端建议的 skill 版本）。`skillLatest` 高于当前 skill 版本时命令会提示更新 skill 目录：https://github.com/ferocknew/vb6-api/tree/main/skill
- `guide` 报 404 = 服务端 DLL < 0.1.17（无此端点）：提示用户更新 DLL（https://github.com/ferocknew/vb6-api 下载后跑 register_dll.vbs 并重启 IDE），勿继续盲写代码。

## 代码写入硬性规则（skill.js 已内置校验的标 ✔，仍需你遵守的标 ✎）

1. ✔ **禁止 `Option Explicit` / `Attribute` 声明头**：skill.js 对写入文件首行做拦截；规则本身——写入内容从过程主体或正文开始，声明头由工程管理（模块级 PUT 会自动保留工程已有 Option 设置）。
2. ✎ **`.SetFocus` 必须写 `Form_Activate()`，禁止写 `Form_Load()`**（否则运行时错误 91）。生成窗体事件代码时遵守。
3. ✔ **换行 CRLF + 尾换行**：skill.js 自动把 LF 归一为 `\r\n` 并补尾换行。
4. ✎ **过程级操作范围含上方紧邻注释与空行**：`proc-put`/`proc-delete` 会连带删除过程上方注释——要保留注释，替换代码里必须自己带上。
5. ✎ **行级写入同样剥声明头且不回插**：模块 Option 设置的修改走 `code-put`（整写自动保留）。

## 典型工作流

1. **诊断/修 bug**：`status` → `snapshot`（或 `code-get` 细看）→ `procs` 定位 → `proc-put` 替换 → `proc-get` 读回逐字符确认。
2. **生成/重写模块**：标准/类模块用 `code-put`（可加 `--type` 创建）；**窗体/用户控件禁止整写**，改用 `proc-put`/过程级操作。
3. **写代码的正确姿势**：先把代码写进临时文件（如 `tmp_code.vb`，UTF-8），再 `code-put <mod> tmp_code.vb`——杜绝 shell 转义与引号问题。
4. **删前必读**：任何 delete 前先 `code-get`/`proc-get` 确认目标确实多余。
5. **编译-修改循环（写码后必做）**：`code-put`/`proc-put` → `compile` 验证 → 失败则按输出「`模块 行号 描述`」逐条修复（行号已换算为 IDE 1 基准，可直接 `code-get`/`proc-get` 对照）→ 再 `compile`，循环直至编译成功。影子编译抓「变量未定义/类型不匹配/语法错误」等编译级问题，与文本级静态审查互补；隐含 save 落盘，无需先手动 `save`。IDE 处于运行态时 compile 拒绝（400），先 `debug-stop`。

## 禁止 / 慎用

- `module-delete` / `proc-delete` 不可逆且无回滚，删前必读确认。
- `code-put` 仅标准/类模块；设计器模块（窗体/用户控件）整写会被服务端拒绝，不要尝试绕过。
- 一次只做一件事：局部改动用过程级命令，不要整模块重写。
- 连续修改时逐个「写→读回验证」，不要写多个后统一验证。
- 修改的是 IDE 内存态：未保存的改动 IDE 崩溃即丢失，重要修改后提醒用户保存工程。

## 环境排查备注（32 位进程）

- VB6 IDE（VB6.EXE）及其编译产物都是 **32 位进程**：排查进程模块/加载 DLL 等信息（如 `Get-Process VB6 | Select -ExpandProperty Modules`）必须用 **32 位 PowerShell**：`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`——64 位 PowerShell 枚举不了 32 位进程的模块。注册表同理查 `WOW6432Node`、regsvr32 用 SysWOW64 版（位数不对齐则查不到/注册不上）。
