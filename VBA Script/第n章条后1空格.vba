'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'作者: 小恐龙'
'日期: 2026年8月9日

Sub FormatArticleSpacing()
    Dim para As Paragraph
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim matchRange As Range
    
    ' 创建正则表达式对象
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = False ' 每段只处理一次
        .IgnoreCase = False
        ' 匹配段落开头的前导空格(组1)、编号(组2)、后续空格(组3)
        .Pattern = "^( *)(第[零一二三四五六七八九十百千\d]+[条章])([ \t]*)"
    End With
    
    ' 遍历文档中的所有段落
    For Each para In ActiveDocument.Paragraphs
        ' 检查段落是否匹配
        Set matches = regEx.Execute(para.Range.Text)
        If matches.Count > 0 Then
            Set match = matches(0)
            ' 构建匹配范围（仅匹配部分，不包括段落标记）
            Set matchRange = para.Range
            matchRange.Start = para.Range.Start + match.FirstIndex
            matchRange.End = matchRange.Start + match.Length
            
            ' 构造新文本：前导空格 + 编号 + 单个空格
            Dim newText As String
            newText = match.SubMatches(0) & match.SubMatches(1) & " "
            
            ' 替换该范围文本（不影响段落其他内容）
            matchRange.Text = newText
        End If
    Next para
    
    Set regEx = Nothing
    Set matchRange = Nothing
    MsgBox "格式调整完成！", vbInformation
End Sub