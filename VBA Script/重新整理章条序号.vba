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


' 主程序：先章后条，各自从1开始顺序编号
Sub 整理章条序号()
    Application.ScreenUpdating = False
    Call ProcessHeading("章")
    Call ProcessHeading("条")
    Application.ScreenUpdating = True
    MsgBox "整理完成！", vbInformation
End Sub

' 通用处理函数：逐查找并替换，只处理段落开头的匹配
Function ProcessHeading(typeChar As String) As Boolean
    Dim rng As Range
    Dim counter As Integer
    counter = 1
    
    Dim pattern As String
    pattern = "第[0-9一二三四五六七八九十]@" & typeChar
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .MatchWildcards = True
        .Text = pattern
        
        Do While .Execute(Forward:=True) = True
            ' 只处理位于段落开头的匹配项（忽略正文引用）
            If .Parent.Start = .Parent.Paragraphs(1).Range.Start Then
                .Parent.Text = "第" & IntegerToChinese(counter) & typeChar
                counter = counter + 1
            End If
            rng.Start = .Parent.End
            rng.End = ActiveDocument.Content.End
        Loop
    End With
    
    ProcessHeading = (counter > 1)
End Function

' 整数→中文数字（支持1~99）
Function IntegerToChinese(num As Integer) As String
    Dim units As Variant
    units = Array("零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    If num <= 10 Then
        IntegerToChinese = units(num)
    ElseIf num < 20 Then
        IntegerToChinese = "十" & units(num - 10)
    ElseIf num Mod 10 = 0 Then
        IntegerToChinese = units(num \ 10) & "十"
    Else
        IntegerToChinese = units(num \ 10) & "十" & units(num Mod 10)
    End If
End Function