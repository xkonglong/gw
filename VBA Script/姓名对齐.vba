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
'日期:2023年2月4日'


Sub xkonglong_所选姓名对齐()
    
    Dim n As Long
    n = InputBox("请确认所选姓名未用空格对齐,再使用本功能." & Chr(13) & "所选姓名几字对齐? 请输入一个整数(默认为三字姓名):", "小恐龙VBA", 3)
    If IsNumeric(n) = False Then
        n = 3
    End If
    
    Dim w As Double
    w = n * Selection.Font.Size
    
    Selection.find.ClearFormatting
    Selection.find.Replacement.ClearFormatting
    With Selection.find
        .Text = " "
        .Replacement.Text = "^p"
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.find.Execute Replace:=wdReplaceAll
    Selection.Range.FitTextWidth = w
End Sub


