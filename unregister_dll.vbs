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
        RunAsAdmin RegSvr, "/u /s """ & fileList(fileName) & """"
        WScript.Sleep 1000
    Next
    Dim wsh
    Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegDelete "HKCU\Software\Microsoft\Visual Basic\6.0\AddIns\InfoAddin.Connect\LoadBehavior"
    wsh.RegDelete "HKCU\Software\Microsoft\Visual Basic\6.0\AddIns\InfoAddin.Connect\CommandLineSafe"
    wsh.RegDelete "HKCU\Software\Microsoft\Visual Basic\6.0\AddIns\InfoAddin.Connect\"
    On Error GoTo 0
    MsgBox "已完成DLL文件的卸载操作！" & vbLF & vbLF & _
           "成功卸载了 " & fileList.Count & " 个文件：" & vbLF & fileNames & vbLF & vbLF & _
           "已从 VB6 外接程序列表移除。", vbInformation, "完成"
End With

Sub RunAsAdmin(program, args)
    CreateObject("Shell.Application").ShellExecute program, args, "", "runas", 1
End Sub
