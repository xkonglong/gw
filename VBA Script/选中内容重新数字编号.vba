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
'日期: 2026年9月8日
'功能：所选内容标题从数字1.开始排序，并删除旧编号。


Sub ReNumberSelection()
    Dim rng As Range
    Dim para As Paragraph
    Dim i As Integer
    Dim text As String
    Dim newText As String
    Dim regex As Object
    
    ' 创建正则表达式对象（匹配开头的数字 + 点或顿号 + 可选空格）
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "^\s*\d+[\.、]\s*"   ' 可匹配 "1. "、"8. "、"1、" 等
    
    ' 如果没有选中任何内容，提示并退出
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选择要重新编号的文本区域。", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection.Range
    i = 1
    
    ' 遍历选中的每个段落
    For Each para In rng.Paragraphs
        text = para.Range.Text
        
        ' 跳过空行（仅包含换行符或空格）
        If Len(Trim(text)) > 1 Then
            ' 如果段落开头有旧编号，则去掉它
            If regex.Test(text) Then
                newText = regex.Replace(text, "")
            Else
                newText = text
            End If
            ' 去除多余首尾空格，并拼接新编号
            newText = Trim(newText)
            para.Range.Text = i & ". " & newText
            i = i + 1
        End If
    Next para
    
    MsgBox "重新编号完成，共处理 " & (i - 1) & " 个条目。", vbInformation
End Sub