# vb6-api —— 给 VB6 IDE 装一个 REST API

VB6 IDE 外接插件（InfoAddin）：在 IDE 进程内开启 HTTP 服务，把当前打开的工程暴露为 REST API——读改代码、管理模块/过程/行、窗体控件、引用、断点、调试。为 AI 编程助手、外部工具、自动化测试提供程序化操控 VB6 IDE 的能力。

## 它能做什么

- 工程/组件/快照只读查询
- 模块级（建/删/整写/导入）、过程级（列/读/改/删，含 Private/Friend 过程）、行级（读/插/改/删）
- 引用库管理（增删、全系统类型库检索）
- 窗体控件（增删改、11 种对齐、窗体样式）
- IDE 工具窗口开关、工程保存
- 调试状态、断点管理、run 启动

共 16 组端点，全部实测回归通过。服务在线后浏览器打开 `http://127.0.0.1:8306/docs` 可看交互式文档（Swagger UI），`/openapi.json` 输出 OpenAPI 规范。

## 系统要求

- Windows（64 位亦可，注册脚本自动改用 SysWOW64 注册 32 位 DLL）
- VB6 IDE（SP6）
- 可选 Node.js ≥ 18（用 skill.js）；没有 node 也没关系，skill.vbs 零依赖（Windows 自带 cscript）

## 安装（一键）

1. 下载本仓库全部文件，保持目录结构（`InfoAddin.dll` 与注册脚本须在同一文件夹）
2. 双击 `register_dll.vbs`（UAC 弹窗点允许）
3. 启动（或重启）VB6 IDE——插件自动加载，服务自动监听 127.0.0.1:8306（端口被占自动 +1 顺延）

验证安装：

```
curl http://127.0.0.1:8306/api/status
```

返回 `{"ok":true,...}` 即成功。

## SKILL 使用

命令行工具在 `skill/` 子文件夹，在仓库根目录执行：

```
node skill/skill.js status          # 或：cscript //nologo skill\skill.vbs status
node skill/skill.js snapshot        # 工程全量快照
node skill/skill.js procs Module1   # 列出模块内全部过程
node skill/skill.js code-get Module1
node skill/skill.js win-show immediate
```

完整命令清单见 [skill/SKILL.md](skill/SKILL.md)；也可以用任意 HTTP 客户端直接调 REST API。

### workBuddy 安装
- 把skill 目录压缩成zip 并导入
<img width="2374" height="810" alt="8809a26e6c94186a6fcdf8a00dd2ba19" src="https://github.com/user-attachments/assets/cdff4ff8-9400-41ed-8a8e-8219afdea981" />

<img width="982" height="726" alt="7eb798395134acee4a4932d66b85178e" src="https://github.com/user-attachments/assets/87cd1272-f5f2-43c1-b121-c01923c0e50f" />

### TRAE 技能导入
- https://docs.trae.ai/ide/skills?_lang=zh

<img width="2336" height="956" alt="image" src="https://github.com/user-attachments/assets/b2b9752c-954e-4ce2-9921-aae4fb914c63" />


## 安全说明

- 服务只监听 127.0.0.1（本机回环），不对外网暴露
- 无鉴权：本机任何进程都可调用，请勿在不受信任的环境使用
- 删除类操作（模块/过程/控件）不可逆，应先读后写

## 卸载

双击 `unregister_dll.vbs`，然后删除本目录。

## 已知边界

- 运行态下 `debug-stop`/`debug-output` 返回 400（防 IDE 挂死的守卫），回设计态需人工点 IDE 停止按钮
- 立即窗口停靠时读不到内容（浮动态可读）
- IDE 里手工 F9 打的断点对 API 不可见（API 只返回自己设置的断点）

## GitHub 发布

- 仓库地址：<https://github.com/ferocknew/vb6-api>
- 获取：`git clone https://github.com/ferocknew/vb6-api.git`，或仓库页绿色 **Code** → **Download ZIP**
- 更新：关闭 VB6 IDE → 用新版覆盖本目录文件 → 双击 `unregister_dll.vbs` 注销旧版 → 双击 `register_dll.vbs` 注册新版 → 重启 IDE（DLL 每次编译 CLSID 都会变，覆盖后必须重新注册）
- 当前仅发布编译产物，是否开源取决于关注量（见下节反馈）

## 反馈

当前仅发布编译产物。如果你有兴趣看到源码开源，或想要新功能，请点 **Star** 或开 **Issue**——关注量将决定开源与后续开发的节奏。

版本 0.1.11（2026-09-06）
