' vb6ide-api skill.vbs — cmd/cscript 版请求工具集（无需 Node.js / PowerShell）
' 用法：cscript //nologo skill.vbs <command> [args]
' 规范与 skill.js 同构：端口探测、非 ASCII 转 \uXXXX、CRLF 归一、信封判定、代码体一律从文件读。
' 依赖：MSXML2.ServerXMLHTTP、Scripting.FileSystemObject（全 Windows 自带，零第三方依赖）。
Option Explicit

Const PORT_BASE = 8306
Const PORT_TRIES = 20
Const TRUNC = 60000

Dim g_port, posArr(32), posCnt, sw, swFull, swMax, swType
Dim optK(31), optV(31), optCnt, swCleanup
g_port = 0
posCnt = 0
sw = ""
swFull = False
swMax = 0
swType = ""
optCnt = 0
swCleanup = False

' 取 --key 参数值（大小写不敏感；无该参数返回空串）
Function Opt(k)
    Dim i
    Opt = ""
    For i = 0 To optCnt - 1
        If StrComp(optK(i), k, vbTextCompare) = 0 Then
            Opt = optV(i)
            Exit Function
        End If
    Next
End Function

' ---------- 工具函数 ----------
Sub Fail(m)
    WScript.Echo "ERROR: " & m
    WScript.Quit 1
End Sub

Sub OutData(s, bFull, nMax)
    Dim limit
    limit = TRUNC
    If nMax > 0 Then limit = nMax
    If Not bFull And Len(s) > limit Then
        s = Left(s, limit) & vbCrLf & "...[截断，共 " & Len(s) & " 字符；--full 全量或 --max N 调整]"
    End If
    WScript.Echo s
End Sub

Function P(n)
    If n < posCnt Then P = posArr(n) Else P = ""
End Function

' JSON 字符串值转义：引号/反斜杠/控制字符 + 非 ASCII 全部 \uXXXX
' （\uXXXX 化同时是服务端 body 编码缺陷的规避手段，与 skill.js 一致）
Function JsonEsc(s)
    Dim i, n, c, k, out
    out = ""
    n = Len(s)
    For i = 1 To n
        c = Mid(s, i, 1)
        k = AscW(c)
        If k < 0 Then k = k + 65536
        Select Case True
            Case c = """": out = out & "\"""
            Case c = "\": out = out & "\\"
            Case k = 13: out = out & "\r"
            Case k = 10: out = out & "\n"
            Case k = 9: out = out & "\t"
            Case k < 32 Or k > 126: out = out & "\u" & Right("0000" & Hex(k), 4)
            Case Else: out = out & c
        End Select
    Next
    JsonEsc = out
End Function

' URL 百分号编码（含非 ASCII 的 UTF-8 序列化）
Function Esc(s)
    Dim out, i2, c2, k2, j2, b0, b1, b2
    out = ""
    For i2 = 1 To Len(s)
        c2 = Mid(s, i2, 1)
        k2 = AscW(c2)
        If k2 < 0 Then k2 = k2 + 65536
        If (k2 >= 48 And k2 <= 57) Or (k2 >= 65 And k2 <= 90) Or (k2 >= 97 And k2 <= 122) Or k2 = 45 Or k2 = 46 Or k2 = 95 Or k2 = 126 Then
            out = out & c2
        ElseIf k2 < 128 Then
            out = out & "%" & Right("0" & Hex(k2), 2)
        ElseIf k2 < 2048 Then
            b0 = 192 Or (k2 \ 64)
            b1 = 128 Or (k2 And 63)
            out = out & "%" & Hex(b0) & "%" & Hex(b1)
        Else
            b0 = 224 Or (k2 \ 4096)
            b1 = 128 Or ((k2 \ 64) And 63)
            b2 = 128 Or (k2 And 63)
            out = out & "%" & Hex(b0) & "%" & Hex(b1) & "%" & Hex(b2)
        End If
    Next
    Esc = out
End Function

' 读代码体文件（UTF-8，兼容 BOM）：剥 BOM、LF 归一 CRLF、补尾换行、拦截声明头（规则 1）
Function ReadCode(f)
    Dim fso, ts, s, first, re
    Set fso = CreateObject("Scripting.FileSystemObject")
    If f = "-" Then
        s = WScript.StdIn.ReadAll
    Else
        If Not fso.FileExists(f) Then Fail "文件不存在：" & f
        Set ts = fso.OpenTextFile(f, 1, False, -1) ' -1 = TristateTrue 读 UTF-8
        s = ts.ReadAll
        ts.Close
    End If
    If Len(s) > 0 Then
        If AscW(Mid(s, 1, 1)) = &HFEFF Then s = Mid(s, 2) ' 去 BOM
    End If
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, vbCrLf)
    If Len(s) > 0 Then
        If Right(s, 2) <> vbCrLf Then s = s & vbCrLf
    End If
    first = Split(s, vbCrLf)(0)
    Set re = New RegExp
    re.Pattern = "^\s*(Option\s+Explicit|Attribute\s+)"
    re.IgnoreCase = True
    If re.Test(first) Then Fail "写入内容首行是 """ & Trim(first) & """：禁止 Option Explicit / Attribute 声明头（规则 1）"
    ReadCode = s
End Function

Function EnvOf(n)
    EnvOf = WScript.CreateObject("WScript.Shell").Environment("PROCESS")(n) & ""
End Function

' 端口探测（8306 起顺延）
Sub FindPort
    Dim p, http, ok
    If Len(EnvOf("VB6IDE_PORT")) > 0 Then
        g_port = CLng(EnvOf("VB6IDE_PORT"))
        Exit Sub
    End If
    For p = PORT_BASE To PORT_BASE + PORT_TRIES - 1
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
        On Error Resume Next
        http.Open "GET", "http://127.0.0.1:" & p & "/api/status", False
        http.SetTimeouts 2000, 2000, 2000, 2000
        http.Send ""
        ok = (Err.Number = 0 And InStr(http.responseText, """ok"":true") > 0)
        On Error GoTo 0
        If ok Then
            g_port = p
            Exit Sub
        End If
    Next
    Fail "服务不可达（" & PORT_BASE & " 起 " & PORT_TRIES & " 个端口无应答）：请打开 VB6 IDE、加载插件并启动服务"
End Sub

' 请求核心：返回响应 JSON 文本；信封 ok:false 时报错退出
Function HttpReq(m, pth, bodyJson)
    Dim http, txt
    FindPort
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    On Error Resume Next
    http.Open m, "http://127.0.0.1:" & g_port & pth, False
    http.SetTimeouts 60000, 60000, 60000, 60000
    If Len(bodyJson) > 0 Then
        http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        http.Send bodyJson
    Else
        http.Send ""
    End If
    If Err.Number <> 0 Then
        On Error GoTo 0
        Fail "请求失败：" & Err.Description
    End If
    On Error GoTo 0
    txt = http.responseText
    If InStr(txt, """ok"":false") > 0 Then
        Fail "HTTP " & http.Status & " ok:false — " & ExtractMsg(txt)
    End If
    HttpReq = txt
End Function

Function ExtractMsg(j)
    Dim re, m
    Set re = New RegExp
    re.Pattern = """message""\s*:\s*""(([^""]|\\"")*)"""
    Set m = re.Execute(j)
    If m.Count > 0 Then ExtractMsg = m(0).SubMatches(0) Else ExtractMsg = j
End Function

Function Qs(q)
    If Len(q) > 0 Then Qs = "?" & q Else Qs = ""
End Function

' 由 --start/--end 参数拼查询串
Function RangeQ()
    Dim q
    q = ""
    If Opt("start") <> "" Then q = "start=" & Opt("start")
    If Opt("end") <> "" Then
        If q <> "" Then q = q & "&"
        q = q & "end=" & Opt("end")
    End If
    RangeQ = Qs(q)
End Function

' 读原始文件（UTF-8 兼 BOM）：JSON 请求体直发用，不做 CRLF 归一与声明头拦截（区别于 ReadCode）
Function ReadRaw(f)
    Dim fso, ts, s
    If f = "-" Then Fail "vbs 版不支持 stdin(-)，请写临时文件后传路径"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(f) Then Fail "文件不存在：" & f
    Set ts = fso.OpenTextFile(f, 1, False, -1)
    s = ts.ReadAll
    ts.Close
    If Len(s) > 0 And AscW(Mid(s, 1, 1)) = 65279 Then s = Mid(s, 2)   ' 剥 BOM
    ReadRaw = s
End Function

' 逗号分隔串 → JSON 字符串数组（[..] 形态，项经 JsonEsc）
Function CsvJsonArr(csv)
    Dim arr, i, s
    arr = Split(csv, ",")
    s = "["
    For i = 0 To UBound(arr)
        If Trim(arr(i)) <> "" Then
            If Len(s) > 1 Then s = s & ","
            s = s & """" & JsonEsc(Trim(arr(i))) & """"
        End If
    Next
    CsvJsonArr = s & "]"
End Function

' ---------- 主流程：参数解析与命令分派 ----------
Dim args, cmd, i, a, curOpt
Set args = WScript.Arguments
If args.Count = 0 Then
    cmd = "help"
Else
    cmd = args(0)
End If
curOpt = ""
For i = 1 To args.Count - 1
    a = args(i)
    If curOpt <> "" Then
        optK(optCnt) = curOpt
        optV(optCnt) = a
        optCnt = optCnt + 1
        curOpt = ""
    ElseIf a = "--cleanup" Then
        swCleanup = True
    ElseIf a = "--full" Then
        swFull = True
    ElseIf Left(a, 2) = "--" And Len(a) > 2 Then
        curOpt = Mid(a, 3)
    Else
        If posCnt <= 31 Then
            posArr(posCnt) = a
            posCnt = posCnt + 1
        End If
    End If
Next
If Opt("max") <> "" Then swMax = CLng(Opt("max"))

Dim tp, bj
Select Case cmd
    Case "help"
        WScript.Echo "vb6ide-api 工具集（cscript //nologo skill.vbs <command>）命令与 skill.js 同构："
        WScript.Echo "  status | project | components | snapshot | ide"
        WScript.Echo "  code-get | module-create | code-put | module-delete | module-import"
        WScript.Echo "  procs | proc-get | proc-put | proc-delete"
        WScript.Echo "  lines-get | lines-insert | lines-replace | lines-delete"
        WScript.Echo "  refs | ref-add | ref-del | typelibs"
        WScript.Echo "  form-get | ctrl-add | ctrl-set | ctrl-del | form-align | form-style"
        WScript.Echo "  win-show | save | debug-status | debug-run | debug-stop"
        WScript.Echo "  bps | bp-set | bp-del | bp-clear | debug-output"
        WScript.Echo "公共项：--full / --max N / 环境变量 VB6IDE_PORT；各命令用法见同名 skill.js"
    Case "status"
        FindPort
        WScript.Echo "port: " & g_port
        OutData HttpReq("GET", "/api/status", ""), swFull, swMax
    Case "project"
        OutData HttpReq("GET", "/api/project", ""), swFull, swMax
    Case "components"
        OutData HttpReq("GET", "/api/components", ""), swFull, swMax
    Case "snapshot"
        OutData HttpReq("GET", "/api/project/snapshot", ""), swFull, swMax
    Case "code-get"
        If P(0) = "" Then Fail "用法：code-get <module> [--start N --end M]"
        OutData HttpReq("GET", "/api/modules/" & Esc(P(0)) & "/code" & RangeQ()), swFull, swMax
    Case "module-create"
        If P(0) = "" Then Fail "用法：module-create <name> [type]"
        tp = "module"
        If P(1) <> "" Then tp = P(1)
        OutData HttpReq("POST", "/api/modules", "{""name"":""" & JsonEsc(P(0)) & """,""type"":""" & JsonEsc(tp) & """}"), swFull, swMax
    Case "code-put"
        If P(0) = "" Or P(1) = "" Then Fail "用法：code-put <module> <file|-> [--type t]"
        bj = "{""code"":""" & JsonEsc(ReadCode(P(1))) & """"
        If Opt("type") <> "" Then bj = bj & ",""type"":""" & Opt("type") & """"
        bj = bj & "}"
        OutData HttpReq("PUT", "/api/modules/" & Esc(P(0)) & "/code", bj), swFull, swMax
    Case "module-delete"
        If P(0) = "" Then Fail "用法：module-delete <name>（不可逆，删前先 code-get 确认）"
        OutData HttpReq("DELETE", "/api/modules/" & Esc(P(0)), ""), swFull, swMax
    Case "module-import"
        If P(0) = "" Then Fail "用法：module-import <绝对路径>"
        OutData HttpReq("POST", "/api/modules/import", "{""path"":""" & JsonEsc(P(0)) & """}"), swFull, swMax
    Case "procs"
        If P(0) = "" Then Fail "用法：procs <module>"
        OutData HttpReq("GET", "/api/modules/" & Esc(P(0)) & "/procedures", ""), swFull, swMax
    Case "proc-get"
        If P(0) = "" Or P(1) = "" Then Fail "用法：proc-get <module> <proc>（Property 加 .Get/.Let/.Set）"
        OutData HttpReq("GET", "/api/modules/" & Esc(P(0)) & "/procedures/" & Esc(P(1)), ""), swFull, swMax
    Case "proc-put"
        If P(0) = "" Or P(1) = "" Or P(2) = "" Then Fail "用法：proc-put <module> <proc> <file|->（注意：操作范围含过程上方注释）"
        OutData HttpReq("PUT", "/api/modules/" & Esc(P(0)) & "/procedures/" & Esc(P(1)), "{""code"":""" & JsonEsc(ReadCode(P(2))) & """}"), swFull, swMax
    Case "proc-delete"
        If P(0) = "" Or P(1) = "" Then Fail "用法：proc-delete <module> <proc> [--cleanup]（不可逆）"
        bj = ""
        If swCleanup Then bj = "cleanup=1"
        OutData HttpReq("DELETE", "/api/modules/" & Esc(P(0)) & "/procedures/" & Esc(P(1)) & Qs(bj)), swFull, swMax
    Case "ide"
        OutData HttpReq("GET", "/api/ide", ""), swFull, swMax
    Case "lines-get"
        If P(0) = "" Then Fail "用法：lines-get <module> [--start N --end M]"
        OutData HttpReq("GET", "/api/modules/" & Esc(P(0)) & "/lines" & RangeQ()), swFull, swMax
    Case "lines-insert"
        If P(0) = "" Or P(1) = "" Or Opt("start") = "" Then Fail "用法：lines-insert <module> --start N <file|->"
        OutData HttpReq("POST", "/api/modules/" & Esc(P(0)) & "/lines", "{""start"":" & Opt("start") & ",""code"":""" & JsonEsc(ReadCode(P(1))) & """}"), swFull, swMax
    Case "lines-replace"
        If P(0) = "" Or P(1) = "" Or Opt("start") = "" Or Opt("end") = "" Then Fail "用法：lines-replace <module> --start N --end M <file|->"
        OutData HttpReq("PATCH", "/api/modules/" & Esc(P(0)) & "/lines", "{""start"":" & Opt("start") & ",""end"":" & Opt("end") & ",""code"":""" & JsonEsc(ReadCode(P(1))) & """}"), swFull, swMax
    Case "lines-delete"
        If P(0) = "" Or Opt("start") = "" Or Opt("end") = "" Then Fail "用法：lines-delete <module> --start N --end M（不可逆）"
        OutData HttpReq("DELETE", "/api/modules/" & Esc(P(0)) & "/lines" & RangeQ()), swFull, swMax
    Case "refs"
        OutData HttpReq("GET", "/api/references", ""), swFull, swMax
    Case "ref-add"
        If Opt("guid") <> "" Then
            tp = Opt("major")
            If tp = "" Then tp = "1"
            bj = Opt("minor")
            If bj = "" Then bj = "0"
            OutData HttpReq("POST", "/api/references", "{""guid"":""" & JsonEsc(Opt("guid")) & """,""major"":" & tp & ",""minor"":" & bj & "}"), swFull, swMax
        ElseIf Opt("path") <> "" Then
            OutData HttpReq("POST", "/api/references", "{""path"":""" & JsonEsc(Opt("path")) & """}"), swFull, swMax
        ElseIf Opt("name") <> "" Then
            OutData HttpReq("POST", "/api/references", "{""name"":""" & JsonEsc(Opt("name")) & """}"), swFull, swMax
        Else
            Fail "用法：ref-add --guid {GUID} [--major N --minor N] | --path 路径 | --name 库名"
        End If
    Case "ref-del"
        If P(0) = "" Then Fail "用法：ref-del <GUID>"
        OutData HttpReq("DELETE", "/api/references/" & Esc(P(0)), ""), swFull, swMax
    Case "typelibs"
        If P(0) <> "" Then
            OutData HttpReq("GET", "/api/typelibs?keyword=" & Esc(P(0)), ""), swFull, swMax
        Else
            OutData HttpReq("GET", "/api/typelibs", ""), swFull, swMax
        End If
    Case "form-get"
        If P(0) = "" Then Fail "用法：form-get <form>"
        OutData HttpReq("GET", "/api/forms/" & Esc(P(0)), ""), swFull, swMax
    Case "ctrl-add"
        If P(0) = "" Or P(1) = "" Or P(2) = "" Then Fail "用法：ctrl-add <form> <name> <className> [--left --top --width --height --caption]"
        bj = "{""name"":""" & JsonEsc(P(1)) & """,""className"":""" & JsonEsc(P(2)) & """"
        If Opt("left") <> "" Then bj = bj & ",""left"":" & Opt("left")
        If Opt("top") <> "" Then bj = bj & ",""top"":" & Opt("top")
        If Opt("width") <> "" Then bj = bj & ",""width"":" & Opt("width")
        If Opt("height") <> "" Then bj = bj & ",""height"":" & Opt("height")
        If Opt("caption") <> "" Then bj = bj & ",""caption"":""" & JsonEsc(Opt("caption")) & """"
        bj = bj & "}"
        OutData HttpReq("POST", "/api/forms/" & Esc(P(0)) & "/controls", bj), swFull, swMax
    Case "ctrl-set"
        If P(0) = "" Or P(1) = "" Or P(2) = "" Then Fail "用法：ctrl-set <form> <ctrl> <props.json>（文件形如 {""props"":{...}}，vbs 版原文直发）"
        OutData HttpReq("PATCH", "/api/forms/" & Esc(P(0)) & "/controls/" & Esc(P(1)), ReadRaw(P(2))), swFull, swMax
    Case "ctrl-del"
        If P(0) = "" Or P(1) = "" Then Fail "用法：ctrl-del <form> <ctrl>（不可逆，删前先 form-get 确认）"
        OutData HttpReq("DELETE", "/api/forms/" & Esc(P(0)) & "/controls/" & Esc(P(1)), ""), swFull, swMax
    Case "form-align"
        If P(0) = "" Or Opt("mode") = "" Or Opt("controls") = "" Then Fail "用法：form-align <form> --mode M --controls a,b,c [--value N]"
        bj = "{""mode"":""" & JsonEsc(Opt("mode")) & """,""controls"":" & CsvJsonArr(Opt("controls"))
        If Opt("value") <> "" Then bj = bj & ",""value"":" & Opt("value")
        bj = bj & "}"
        OutData HttpReq("POST", "/api/forms/" & Esc(P(0)) & "/align", bj), swFull, swMax
    Case "form-style"
        If P(0) = "" Or P(1) = "" Then Fail "用法：form-style <form> <style.json>（文件即请求体原文）"
        OutData HttpReq("PATCH", "/api/forms/" & Esc(P(0)) & "/style", ReadRaw(P(1))), swFull, swMax
    Case "win-show"
        tp = P(0)
        If tp = "" Then tp = "immediate"
        OutData HttpReq("POST", "/api/ui/immediate-window", "{""window"":""" & JsonEsc(tp) & """}"), swFull, swMax
    Case "save"
        OutData HttpReq("POST", "/api/ui/save-project", "{}"), swFull, swMax
    Case "debug-status"
        OutData HttpReq("GET", "/api/debug/status", ""), swFull, swMax
    Case "debug-run"
        OutData HttpReq("POST", "/api/debug/run", "{}"), swFull, swMax
    Case "debug-stop"
        OutData HttpReq("POST", "/api/debug/stop", "{}"), swFull, swMax
    Case "bps"
        OutData HttpReq("GET", "/api/debug/breakpoints", ""), swFull, swMax
    Case "bp-set"
        If P(0) = "" Or P(1) = "" Then Fail "用法：bp-set <module> <line>（toggle：该行已有断点则本次为取消）"
        OutData HttpReq("POST", "/api/debug/breakpoints", "{""module"":""" & JsonEsc(P(0)) & """,""line"":" & P(1) & "}"), swFull, swMax
    Case "bp-del"
        If P(0) <> "" And P(1) <> "" Then
            OutData HttpReq("DELETE", "/api/debug/breakpoints", "{""module"":""" & JsonEsc(P(0)) & """,""line"":" & P(1) & "}"), swFull, swMax
        Else
            OutData HttpReq("DELETE", "/api/debug/breakpoints", "{}"), swFull, swMax
        End If
    Case "bp-clear"
        OutData HttpReq("DELETE", "/api/debug/breakpoints", "{}"), swFull, swMax
    Case "debug-output"
        OutData HttpReq("GET", "/api/debug/output", ""), swFull, swMax
    Case Else
        Fail "未知命令：" & cmd & "（skill.vbs help 查看清单）"
End Select
