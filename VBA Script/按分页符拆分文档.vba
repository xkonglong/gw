'====================================================================
'                            使用说明:
'
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'  作者: 小恐龙 (改编)
'  日期: 2026年8月31日
'  说明: 根据手动分页符拆分文档，每个分页符处拆分为一个新文档。
'====================================================================

Sub User_根据分页符拆分文档()

    ' 检查文档中是否存在手动分页符
    Dim rngFind As Range
    Set rngFind = ActiveDocument.Content
    With rngFind.Find
        .Text = "^m"                ' 手动分页符
        .Forward = True
        .Wrap = wdFindStop
    End With
    If Not rngFind.Find.Execute Then
        MsgBox "当前文档中没有手动分页符，请先插入分页符再运行此宏。"
        Exit Sub
    End If

    ' 选择保存拆分后文档的文件夹
    Dim dia As FileDialog
    Set dia = Application.FileDialog(msoFileDialogFolderPicker)
    dia.Title = "请选择拆分后要保存的文件夹"
    If ActiveDocument.Saved Then dia.InitialFileName = ActiveDocument.Path
    If dia.Show = 0 Then Exit Sub
    Dim sPath As String
    sPath = dia.SelectedItems(1)

    Dim oDoc As Document
    Set oDoc = ActiveDocument
    Dim startRange As Range
    Set startRange = oDoc.Range(0, 0)   ' 起始位置
    Dim findRange As Range
    Dim i As Long
    i = 1
    Dim nDoc As Document
    Dim copyRange As Range

    ' 循环查找所有手动分页符
    Do
        ' 从当前起始位置开始查找下一个分页符
        Set findRange = oDoc.Range(startRange.End, oDoc.Range.End)
        With findRange.Find
            .Text = "^m"
            .Forward = True
            .Wrap = wdFindStop
        End With

        If findRange.Find.Execute Then
            ' 分页符位于 findRange.Start 处
            ' 复制从起始位置到分页符之前的内容（不包括分页符）
            Set copyRange = oDoc.Range(startRange.Start, findRange.Start)
            ' 如果复制区域非空，则创建新文档
            If copyRange.Start < copyRange.End Then
                copyRange.Copy
                Set nDoc = Documents.Add
                nDoc.Content.PasteAndFormat (wdFormatOriginalFormatting)
                ' 复制原文档的页面设置（使用第一节的设置）
                nDoc.PageSetup = oDoc.Sections(1).PageSetup
                ' 保存新文档
                Dim baseName As String
                baseName = Left(oDoc.Name, InStrRev(oDoc.Name, ".") - 1)
                If baseName = "" Then baseName = "未命名文档"
                nDoc.SaveAs2 sPath & "\" & baseName & "_" & i & ".docx"
                nDoc.Close
                i = i + 1
            End If
            ' 更新起始位置为分页符之后
            startRange.Start = findRange.End
        Else
            ' 没有找到更多分页符，退出循环
            Exit Do
        End If
    Loop

    ' 处理剩余部分（从最后起始位置到文档末尾）
    If startRange.Start < oDoc.Range.End Then
        Set copyRange = oDoc.Range(startRange.Start, oDoc.Range.End)
        copyRange.Copy
        Set nDoc = Documents.Add
        nDoc.Content.PasteAndFormat (wdFormatOriginalFormatting)
        nDoc.PageSetup = oDoc.Sections(1).PageSetup
        Dim baseName2 As String
        baseName2 = Left(oDoc.Name, InStrRev(oDoc.Name, ".") - 1)
        If baseName2 = "" Then baseName2 = "未命名文档"
        nDoc.SaveAs2 sPath & "\" & baseName2 & "_" & i & ".docx"
        nDoc.Close
        i = i + 1
    End If

    MsgBox "拆分完成，共生成 " & (i - 1) & " 个文档。", vbInformation
End Sub