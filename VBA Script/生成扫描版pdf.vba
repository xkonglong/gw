'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'作者: 小恐龙
'日期: 2026年7月14日
'功能：直接将word文档转为扫描版pdf,需要用户自行下载 make-look-scanned,并修改112行代码来指定程序路径.
'下载地址: https://github.com/overflowy/make-look-scanned


Sub ConvertCurrentDocToScannedPDF()
    On Error GoTo ErrorHandler

    Dim docPath As String, docDir As String, docBaseName As String
    Dim tempPDF As String, outputPDF As String
    Dim exePath As String
    Dim wsh As Object
    Dim ret As Long
    Dim skew As String, noise As String, grayscale As String, paperTone As String
    Dim blur As String, edgeShadow As String, jpegQuality As String, dpi As String
    Dim cmd As String
    Dim saveAsDialog As FileDialog

    ' ----- 1. 确保当前文档已保存 -----
    If ActiveDocument.Path = "" Then
        MsgBox "当前文档尚未保存，请先保存文档再运行此宏。", vbExclamation
        ' 可选：调用保存对话框让用户保存
        If MsgBox("是否立即保存文档？", vbYesNo + vbQuestion) = vbYes Then
            Application.Dialogs(wdDialogFileSaveAs).Show
            If ActiveDocument.Path = "" Then
                MsgBox "文档未保存，操作取消。", vbCritical
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' 获取文档信息
    docPath = ActiveDocument.FullName
    docDir = ActiveDocument.Path
    docBaseName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)

    ' ----- 2. 用户输入各参数（默认值保持不变）-----
    skew = InputBox("请输入倾斜度 (skew)，默认 0.6：", "参数设置", "0.6")
    If skew = "" Then Exit Sub

    noise = InputBox("请输入噪点强度 (noise)，默认 0.08：", "参数设置", "0.08")
    If noise = "" Then Exit Sub

    grayscale = InputBox("是否灰度化 (grayscale) true / false，默认 false：", "参数设置", "false")
    If grayscale = "" Then Exit Sub

    paperTone = InputBox("请输入纸张色调 (paper-tone)，默认 0：", "参数设置", "0")
    If paperTone = "" Then Exit Sub

    blur = InputBox("请输入模糊程度 (blur)，默认 0.1：", "参数设置", "0.1")
    If blur = "" Then Exit Sub

    edgeShadow = InputBox("请输入边缘阴影 (edge-shadow)，默认 0.05：", "参数设置", "0.05")
    If edgeShadow = "" Then Exit Sub

    jpegQuality = InputBox("请输入 JPEG 质量 (jpeg-quality)，默认 80：", "参数设置", "80")
    If jpegQuality = "" Then Exit Sub

    dpi = InputBox("请输入分辨率 (dpi)，默认 300：", "参数设置", "300")
    If dpi = "" Then Exit Sub

    ' ----- 3. 指定输出 PDF 路径（可自定义，留空则自动生成）-----
    outputPDF = InputBox("请输入输出扫描版 PDF 的完整路径（留空则自动生成）：", "输出设置", "")
    If outputPDF = "" Then
        outputPDF = docDir & "\" & docBaseName & "_scanned.pdf"
    End If
    ' 确保扩展名为 .pdf
    If LCase(Right(outputPDF, 4)) <> ".pdf" Then outputPDF = outputPDF & ".pdf"

    ' ----- 4. 临时 PDF 路径（用于导出原始文档）-----
    tempPDF = docDir & "\~temp_" & docBaseName & ".pdf"

    ' ----- 5. 导出当前文档为临时 PDF -----
    On Error Resume Next
    Kill tempPDF          ' 删除可能存在的旧临时文件
    On Error GoTo ErrorHandler

    ActiveDocument.ExportAsFixedFormat _
        OutputFileName:=tempPDF, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, _
        From:=1, To:=1, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False

    If Dir(tempPDF) = "" Then
        MsgBox "导出 PDF 失败，请检查文档或权限。", vbCritical
        Exit Sub
    End If

    ' ----- 6. 确定 make-look-scanned.exe 位置 -----
    '    优先使用当前文档所在目录，若不存在则让用户手动选择
    exePath = "D:\Tools\Office\PDF工具\文字pdf变扫描\make-look-scanned.exe"
    If Dir(exePath) = "" Then
        MsgBox "本功能需要用户下载安装 make-look-scanned.exe，请手动指定该程序位置。", vbInformation
        Dim fdExe As FileDialog
        Set fdExe = Application.FileDialog(msoFileDialogFilePicker)
        With fdExe
            .Filters.Clear
            .Filters.Add "可执行文件", "*.exe"
            .Title = "请选择 make-look-scanned.exe"
            If .Show = -1 Then
                exePath = .SelectedItems(1)
            Else
                MsgBox "未选择 exe，操作取消。", vbExclamation
                GoTo CleanUp
            End If
        End With
    End If

    ' ----- 7. 构建命令行（包含 --force 和 -o 输出）-----
    cmd = """" & exePath & """ --skew " & skew & _
          " --noise " & noise & _
          " --grayscale=" & grayscale & _
          " --paper-tone " & paperTone & _
          " --blur " & blur & _
          " --edge-shadow " & edgeShadow & _
          " --jpeg-quality " & jpegQuality & _
          " --dpi " & dpi & _
          " --force """ & tempPDF & """ -o """ & outputPDF & """"

    ' ----- 8. 执行转换（等待完成）-----
    Set wsh = CreateObject("WScript.Shell")
    ret = wsh.Run(cmd, 0, True)   ' 0 隐藏窗口，True 等待结束

    If ret <> 0 Then
        MsgBox "转换失败，错误代码：" & ret & vbCrLf & "请检查参数或文件。", vbCritical
        GoTo CleanUp
    End If

    ' ----- 9. 检查输出文件并打开 -----
    If Dir(outputPDF) <> "" Then
        wsh.Run """" & outputPDF & """", 1, False
        MsgBox "转换完成！已打开生成的扫描版 PDF。", vbInformation
    Else
        MsgBox "转换完成，但未找到输出文件：" & outputPDF & vbCrLf & "请手动检查。", vbWarning
    End If

CleanUp:
    ' ----- 10. 删除临时 PDF -----
    On Error Resume Next
    If Dir(tempPDF) <> "" Then Kill tempPDF
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "发生错误：" & vbCrLf & Err.Description, vbCritical
    Resume CleanUp
End Sub