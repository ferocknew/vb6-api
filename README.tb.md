# tb-api —— 给 twinBASIC IDE 装一个 REST API

twinBASIC IDE 插件（tb-info-addin）：在 IDE 进程内开启 HTTP 服务，把当前打开的工程暴露为 REST API——读改代码、行级/过程级编辑、工程快照、表达式求值。为 AI 编程助手、外部工具、自动化测试提供程序化操控 twinBASIC IDE 的能力。是本仓库 VB6 版（[README.md](README.md)）的 twinBASIC 姊妹版，同端口同信封约定。

## 它能做什么

- 工程/模块只读查询：工程元数据、模块清单、代码读取（支持行区间）、一键全量快照
- 代码写通道三粒度：模块整写（PUT code）、行级三操作（插入/替换/删除）、过程级四端点（列/读/替换/删除，文本扫描口径）
- 引用库查询、工程保存、twinBASIC 编码规范动态注入（18 条，含 tb 相对 VB6 的语法差异）
- Debug Console 语境表达式求值（`/api/eval`，VB6 版没有的增量能力）
- Swagger UI 交互式文档 + OpenAPI 3.0 规范输出

共 17 路径 22 操作，38 用例黑盒实测全绿（0.1.11 定版）。

## 系统要求

- Windows
- twinBASIC IDE（BETA 983 及以上；免费 Personal 版即可）

## 安装

1. 下载本仓库
2. 把 `tb-info-addin.dll` 拷贝到 twinBASIC 安装目录下的 `addins\win32\`（例如 `C:\Tools\twinBASIC_IDE_BETA_983\addins\win32\`）
3. 启动（或重启）twinBASIC IDE——插件自动加载，工具栏出现 tb-api 按钮，服务自动监听 127.0.0.1:8306（端口被占自动 +1 顺延）

无需 COM 注册（regsvr32）、无需管理员权限——twinBASIC 插件从 addins 目录自动发现，拷文件即用。

验证安装：

```
curl http://127.0.0.1:8306/api/status
```

返回 `{"ok":true,...,"service":{"version":"0.1.11",...}}` 即成功。浏览器打开 `http://127.0.0.1:8306/docs` 可看交互式文档（Swagger UI）。

## 快速上手

```
curl http://127.0.0.1:8306/api/modules                    # 模块清单
curl http://127.0.0.1:8306/api/modules/MainModule.twin/code   # 读代码
curl http://127.0.0.1:8306/api/project/snapshot            # 全量快照
curl "http://127.0.0.1:8306/api/eval?expr=1+1"             # 表达式求值
```

写操作（body 为 UTF-8 代码文本）：

```
curl -X PATCH "http://127.0.0.1:8306/api/modules/MainModule.twin/lines?start=1&end=1" -d "' 注释替换"
curl -X PUT "http://127.0.0.1:8306/api/modules/MainModule.twin/code" --data-binary @新全文.twin
```

完整端点清单看 `/docs` 或 `/api/openapi.json`。

## 与 VB6 版差异

| 能力 | VB6 版 | tb 版 |
| --- | --- | --- |
| 写通道 | VBE CodeModule 直写 | 编辑器通道整写（tb VFS 只读，Open→Text→Save→Close） |
| 表达式求值 | 无 | ✅ /api/eval |
| 窗体控件操作 | ✅ 设计态对象模型 | ❌（UIDESIGNER 为 JSON 文件，无对象 API） |
| 调试控制/断点 | ✅ | ❌（扩展 API 未提供） |
| 模块新建/删除 | ✅ | ❌（VFS 写侧未实现） |
| 编译验证 | 影子编译（第二进程异步 job） | IDE 内 Build 同步触发（结果查询端点规划中） |
| 安装方式 | regsvr32 注册 + HKCU | 拷 DLL 到 addins 即用 |
| 依赖 | msvbvm60.dll | 无额外运行时 |

## 版本记录

- **0.1.11**（2026-09-09）：修复 PUT 整写假成功（内容未落盘）；38 用例实测全绿定版
- 0.1.10（2026-09-09）：修复 code 读/写通道全 404（路由双重剥前缀）；写通道尾 CRLF 字节保真
- 0.1.9（2026-09-09）：写通道核心（整写/行级/过程级）、工程组端点接线、Swagger UI tags 分组
- 更早：0.1.4-0.1.8 传输层与读端点演进（详见主仓库提交历史）

## 已知限制

- 模块名暂不支持中文/特殊字符（路径不做 URL 解码，ASCII 文件名）
- 写操作全局串行（并发写请求回 409 Conflict）
- `GET /api/build` 触发编译后暂无结果查询端点（可观测性开发中）
- `/api/modules` 列表包含 Resources 下的非代码文件（过滤待优化）
- 写后回读建议延时 ≥500ms（IDE 编辑器落盘有约 120ms 陈旧窗口）
