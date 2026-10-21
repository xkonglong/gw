Sub 中文数字重新编号
    ' 功能：将选定段落中首组中文括号编号（如（一）、（二）…）重新从（一）开始连续编号
    ' 支持编号范围 1~99（如（一）~（九十九））
    ' 仅处理含有编号的段落，未含编号的段落原样保留，编号顺序按段落先后递增

    Dim rng As Range
    Dim para As Paragraph
    Dim i As Integer
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object

    ' 检查是否有选中内容
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选中要重新编号的内容。", vbExclamation, "提示"
        Exit Sub
    End If

    Set rng = Selection.Range
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False          ' 仅匹配每段第一个编号
    regEx.Pattern = "（[一二三四五六七八九十]+）"   ' 匹配中文括号内中文数字（一~十、十一等）

    i = 1
    ' 遍历选中区域内的所有段落
    For Each para In rng.Paragraphs
        Set matches = regEx.Execute(para.Range.Text)
        If matches.Count > 0 Then
            Set match = matches(0)
            
            ' 计算匹配文本在段落中的位置
            Dim matchRange As Range
            Set matchRange = para.Range
            matchRange.Start = para.Range.Start + match.FirstIndex
            matchRange.End = matchRange.Start + match.Length
            
            ' 替换为新的连续编号（调用转换函数）
            matchRange.Text = "（" & 数字转中文(i) & "）"
            i = i + 1
        End If
    Next para

    ' 清理对象
    Set regEx = Nothing
    Set matchRange = Nothing
    Set rng = Nothing

    MsgBox "重新编号完成！共处理 " & (i - 1) & " 个编号。", vbInformation, "完成"
End Sub

' ------------------------------------------------------------
' 辅助函数：将整数（1~99）转换为中文数字（一~九十九）
' ------------------------------------------------------------
Function 数字转中文(ByVal num As Integer) As String
    Dim digits As String
    digits = "一二三四五六七八九"
    
    If num <= 0 Then
        数字转中文 = ""
        Exit Function
    End If
    
    If num <= 9 Then
        数字转中文 = Mid(digits, num, 1)
    ElseIf num <= 99 Then
        Dim shi As Integer, ge As Integer
        shi = num \ 10
        ge = num Mod 10
        If shi = 1 Then
            数字转中文 = "十"
        Else
            数字转中文 = Mid(digits, shi, 1) & "十"
        End If
        If ge > 0 Then
            数字转中文 = 数字转中文 & Mid(digits, ge, 1)
        End If
    Else
        ' 若编号超过99，简单显示为阿拉伯数字（可自行扩展）
        数字转中文 = "(" & num & ")"   ' 此处可改为其他处理
    End If
End Function