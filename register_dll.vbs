'本文件含中文：必须保持 ANSI(GBK) 编码 + CRLF 换行，禁止转为 UTF-8 或 LF（改文件先 iconv 到 UTF-8，改完转回 GBK 并统一 CRLF）

With CreateObject("Scripting.FileSystemObject")
    Dim currentPath: currentPath = .GetParentFolderName(WScript.ScriptFullName)
    Dim RegSvr
    If .FileExists(.GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe") Then
        RegSvr = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe"
    ElseIf .FileExists(.GetSpecialFolder(0).Path & "\System32\regsvr32.exe") Then
        RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe"
    Else
        MsgBox "无法找到regsvr32.exe！", vbCritical, "错误"
        WScript.Quit
    End If
    Dim dllFiles(0)
    dllFiles(0) = "InfoAddin.dll"
    Dim fileList
    Set fileList = CreateObject("Scripting.Dictionary")
    Dim i, missingFiles
    missingFiles = ""
    For i = LBound(dllFiles) To UBound(dllFiles)
        Dim fileName: fileName = dllFiles(i)
        Dim filePath: filePath = currentPath & "\" & fileName
        If Not .FileExists(filePath) Then filePath = currentPath & "\..\src\" & fileName
        If .FileExists(filePath) Then
            fileList.Add fileName, filePath
        Else
            missingFiles = missingFiles & fileName & vbLF
        End If
    Next
    If fileList.Count = 0 Then
        MsgBox "未找到任何DLL文件！" & vbLF & vbLF & "查找路径：" & currentPath, vbExclamation, "错误"
        WScript.Quit
    End If
    Dim fileNames: fileNames = Join(fileList.Keys, vbLF)
    For Each fileName In fileList.Keys
        RunAsAdmin RegSvr, "/s """ & fileList(fileName) & """"
        WScript.Sleep 1000
    Next
    Dim wsh
    Set wsh = CreateObject("WScript.Shell")
    Dim addinKey
    addinKey = "HKCU\Software\Microsoft\Visual Basic\6.0\AddIns\InfoAddin.Connect\"
    On Error Resume Next
    wsh.RegWrite addinKey & "LoadBehavior", 3, "REG_DWORD"
    wsh.RegWrite addinKey & "CommandLineSafe", 0, "REG_DWORD"
    On Error GoTo 0
    ' 坑 31：RegWrite 可能静默失败（实测残留 0x5）。回读 LoadBehavior 校验，不符则报错退出非 0
    On Error Resume Next
    Dim lbVal, lbOk
    lbOk = False
    lbVal = wsh.RegRead(addinKey & "LoadBehavior")
    If Err.Number = 0 Then
        If CLng(lbVal) = 3 Then lbOk = True
    End If
    On Error GoTo 0
    If Not lbOk Then
        Dim msg
        msg = "注册表回读校验失败：LoadBehavior 期望 3，实际=" & lbVal & _
              "（RegWrite 可能静默失败，坑 31）。请手动执行：" & vbLF & _
              "reg add """ & Left(addinKey, Len(addinKey) - 1) & """ /v LoadBehavior /t REG_DWORD /d 3 /f"
        WScript.Echo "ERROR: " & msg
        MsgBox msg, vbCritical, "注册失败"
        WScript.Quit 1
    End If
    MsgBox "已完成DLL文件的注册操作！" & vbLF & vbLF & _
           "已注册 COM 并登记到 VB6 外接程序列表(LoadBehavior=3)。" & vbLF & vbLF & _
           "请重启 VB6 IDE，在外接程序管理器勾选 InfoAddin。", vbInformation, "完成"
End With

Sub RunAsAdmin(program, args)
    CreateObject("Shell.Application").ShellExecute program, args, "", "runas", 1
End Sub
