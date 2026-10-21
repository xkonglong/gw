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

Sub 加粗第n章第n条()
    Application.ScreenUpdating = False
    
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        ' 匹配“第” + 数字/中文数字（1个或多个） + “章”或“条”
        .Text = "第[0-9一二三四五六七八九十]@[章条]"
        
        ' 循环查找所有匹配项
        Do While .Execute(Forward:=True) = True
            ' 获取匹配所在段落
            Dim para As Paragraph
            Set para = .Parent.Paragraphs(1)
            ' 将整个段落加粗
            para.Range.Font.Bold = True
            ' 将查找范围移到该段落后，避免重复处理同一段落
            .Parent.Start = para.Range.End
            .Parent.End = ActiveDocument.Content.End
        Loop
    End With
    
    Application.ScreenUpdating = True
    MsgBox "加粗完成！", vbInformation
End Sub