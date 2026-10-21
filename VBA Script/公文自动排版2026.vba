'====================================================================
'                            使用说明:
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================

'  原作者: 413191246se
'  源码地址: https://club.excelhome.net/thread-1723703-1-1.html
'  修改人: 小恐龙
'  修改更新日期: 2026年9月26日
'  修改说明:  本代码主要删除了与自动排版无关的功能,并对多数关键代码进行了中文注释,方便用户根据情况自行修改

 '------------------------------------------------------------------------
 '文档尽量不要过于复杂, 尤其不推荐图文排版使用本脚本
 '------------------------------------------------------------------------
Option Explicit

' 主过程：公文自动排版
Sub 公文自动排版()
' 依次执行各排版步骤
    PaperA4          ' 设置纸张为A4，页边距等
    Initial          ' 初始化文档：删除域、转换编号、替换特殊字符等
    Title2345        ' 处理标题2~5级（一、/（一）/1．/（1）等）
    Title1           ' 处理一级标题（主标题）
    Inscribe         ' 处理落款（日期、单位、附件等）
    PageNumber       ' 插入页码（公文格式：一字线居中）
    NumPages         ' 根据页数调整显示视图
    Selection.HomeKey Unit:=wdStory  ' 光标移到文档开头
    ActiveDocument.Content.Font.Color = wdColorAutomatic ' 恢复字体颜色为自动
    Options.ReplaceSelection = True '恢复键入时替换所选内容为勾选状态
End Sub

' 设置A4纸张和页边距（兼容横纵向）
Function PaperA4()
    Dim s As Section
    For Each s In ActiveDocument.Sections
        With s.PageSetup
			If .Orientation = wdOrientPortrait Then                      ' 如果页面方向为纵向
				.TopMargin = CentimetersToPoints(2.54)                   ' 上边距 2.54 厘米
				.BottomMargin = CentimetersToPoints(2.54)                ' 下边距 2.54 厘米
				.LeftMargin = CentimetersToPoints(3.17)                  ' 左边距 3.17 厘米（通常留装订位）
				.RightMargin = CentimetersToPoints(3.17)                 ' 右边距 3.17 厘米
				.PageWidth = CentimetersToPoints(21)                     ' 页面宽度 21 厘米（A4纵向宽）
				.PageHeight = CentimetersToPoints(29.7)                  ' 页面高度 29.7 厘米（A4纵向高）
			Else                                                         ' 否则（横向）
				.TopMargin = CentimetersToPoints(2.54)                   ' 上边距 2.54 厘米
				.BottomMargin = CentimetersToPoints(2.54)                ' 下边距 2.54 厘米
				.LeftMargin = CentimetersToPoints(2.54)                  ' 左边距 2.54 厘米（对称）
				.RightMargin = CentimetersToPoints(2.54)                 ' 右边距 2.54 厘米
				.PageWidth = CentimetersToPoints(29.7)                   ' 页面宽度 29.7 厘米（A4横向宽）
				.PageHeight = CentimetersToPoints(21)                    ' 页面高度 21 厘米（A4横向高）
			End If                                                       ' 结束条件判断
            .HeaderDistance = CentimetersToPoints(1.5)   ' 页眉距离
            .FooterDistance = CentimetersToPoints(1.5)   ' 页脚距离
        End With
    Next
End Function

' 初始化文档：删除域、转换编号、替换特殊字符、处理表格环绕等
Function Initial()
    Options.ReplaceSelection = False
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit

    With ActiveDocument
        ' 删除所有域（如自动编号域），转为普通文本
        .Fields.Unlink

        ' 将列表编号转换为普通文本
        .ConvertNumbersToText

        ' 使用查找替换进行一系列格式预处理
        With .Content.Find
            ' 手动换行符 -> 段落标记
            .Execute "^l", , , 0, , , , , , "^p", 2
            ' 删除标点符号后的空格（如“。）”后的空格）
            .Execute "([、.．）)])([ 　^s^t]{1,})", , , 1, , , , , , "\1", 2
            ' 处理“附件”后可能缺少“件”字的情况
            .Execute "(^13附)([!一-﨩])", , , 1, , , , , , "\1件\2", 2
            ' “附表”替换为“附件”
            .Execute "附表", , , , , , , , , "附件", 2
            ' 将年份和文号中的括号改为六角括号〔〕
            .Execute "?([0-9]{4})?([0-9]@号)", , , 1, , , , , , "〔\1〕\2", 2
            ' 将句号后的括号移到句号之前（如“。（1）”变成“（1）。”）
            .Execute "(。)([（(]*[)）])(^13)", , , 1, , , , , , "\2\1\3", 2
        End With

        ' 取消表格的文字环绕，并居中对齐
        Dim t As Table
        For Each t In .Tables
            With t.Rows
                .WrapAroundText = False   ' 取消环绕
                .Alignment = wdAlignRowCenter  ' 表格行居中
            End With
        Next
    End With
    ComeUnderLine   ' 处理下划线占位符
End Function

' 将文档中所有带下划线的空格临时替换为特殊字符“`”，以便后续恢复
Function ComeUnderLine()
    With ActiveDocument.Content.Find
        .Font.Underline = wdUnderlineSingle
        .Execute FindText:=" ", ReplaceWith:="`", Format:=True, Replace:=wdReplaceAll
    End With
End Function

' 处理标题2~5级（一、 (一) 1． (1) 等），设置样式、字体颜色、自动编号等
Function Title2345()
    Const s As String = "[一二三四五六七八九十百千零〇○Oo0０Ｏｏ]@"   ' 匹配中文数字
    Const t As String = "*[。：；，、！？…—.:;,!?]?"              ' 匹配句尾标点
    Const u As String = "[一二三四五六七八九十]"                  ' 匹配单个中文数字
    Dim doc As Document, i As Paragraph, r(), n&, b&, c&, d&, e&
    Dim v2&, p2&, j2&, v3&, p3&, j3&, v4&, p4&, j4&, v5&, p5&, j5&, z&, y&, x&, v6&, p6&

    Set doc = ActiveDocument

    ' 跳过表格区域，避免表格内容被误处理（代码来自sylun）
    With doc
        ReDim r(.Tables.Count + 1)

        If .Tables.Count = 0 Then
            Set r(1) = .Content
        Else
            For n = 1 To .Tables.Count
                If n = 1 Then
                    Set r(n) = .Range(0, .Tables(n).Range.Start)
                Else
                    Set r(n) = .Range(.Tables(n - 1).Range.End, .Tables(n).Range.Start)
                End If
            Next
            Set r(n) = .Range(.Tables(n - 1).Range.End, .Content.End)
        End If
    End With

    For n = 1 To UBound(r)
        With r(n)
            .Select

            ' 删除段落首尾空格
            CommandBars.FindControl(ID:=122).Execute

            ' 清除格式
            Selection.ClearFormatting

            ' 应用正文样式（宋体10.5pt等）
            Word2003NormalStyle

            ' 默认字体设为仿宋，蓝色（临时）
            With .Font
                .Name = "仿宋"
                .Color = wdColorBlue
            End With

            .ParagraphFormat.SpaceAfter = 0   ' 段后间距0

            ' 如果区域不是文档开头，在前面插入一个空段落
            If .Start <> 0 Then .InsertParagraphBefore

            ' 规范化编号前缀：例如将“一、”修正为“一、”，将“（一）”统一格式等
            With .Find
                .Execute "(^13)(" & s & ")(、)", , , 1, , , , , , "\1一\3", 2
                .Execute "(^13)([(（]" & s & "[）)])", , , 1, , , , , , "\1（一）", 2
                .Execute "(^13)([0-9０-９]@[、.．])", , , 1, , , , , , "\11．", 2
                .Execute "(^13)[(（][0-9０-９]@[）)]", , , 1, , , , , , "\1（1）", 2
            End With

            ' 删除空段落（只有回车符的段落）
            For Each i In .Paragraphs
                With i.Range
                    If Asc(.Text) = 13 Then .Delete
                End With
            Next

            ' 特殊处理：如果是空区域或只有一个空段落，则跳过后续处理
            If .Start = 0 Then
                If .Paragraphs.Count = 1 Then GoTo RangeFormat
            Else
                If .Paragraphs.Count = 1 Then
                    If .Text = vbCr Then
                        .Font.Size = 1
                        .ParagraphFormat.LineSpacing = LinesToPoints(0.06)
                        .Select
                        CancelGrid
                        Exit Function
                    End If
                ElseIf .Paragraphs.Count = 2 Then
                    If .Paragraphs(2).Range.End = doc.Content.End Then
                        y = 1
                        GoTo RangeFormat
                    End If
                End If
            End If

            ' 遍历每个段落，识别标题级别并设置样式、重新编号
            For Each i In .Paragraphs
                With i.Range
' ----- 二级标题：一、 -----
                    If .Text Like "一、*" Then
                        v3 = 0
                        v4 = 0

                        If v2 = 0 Then
                            v2 = 1
                            If .Next(4, 1) Like "一、*" Then
                                p2 = 0
                                If .Sentences.Count >= 2 And .Next(4, 1).Sentences.Count >= 2 Then j2 = 2 Else j2 = 1
                            Else
                                p2 = 1
                            End If
                        End If

                        If p2 = 1 Then
                            .Style = wdStyleHeading2
                            With .Font
                                .NameFarEast = "黑体"
                                .NameAscii = "Arial"
                                .Color = wdColorRed
                            End With
                        Else
                            .ParagraphFormat.OutlineLevel = wdOutlineLevel2
                        End If

                        b = b + 1
                        c = 0
                        d = 0
                        e = 0
                        ' 将编号改为自动域（中文数字）
                        With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "、")).Start)
                            .Delete
                            .Fields.Add Range:=.Paragraphs(1).Range, Text:="= " & b & " \* CHINESENUM3"
                            .Fields.Unlink
                        End With
' ----- 三级标题：（一） -----
                    ElseIf .Text Like "（一）*" Then
                        v4 = 0

                        If v3 = 0 Then
                            v3 = 1
                            If .Next(4, 1) Like "（一）*" Then
                                p3 = 0
                                If .Sentences.Count >= 2 And .Next(4, 1).Sentences.Count >= 2 Then j3 = 2 Else j3 = 1
                            Else
                                p3 = 1
                            End If
                        End If

                        If p3 = 1 Then
                            .Style = wdStyleHeading3
                            With .Font
                                .NameFarEast = "楷体"
                                .NameAscii = "Times New Roman"
                                .Color = wdColorPink
                            End With
                        Else
                            .ParagraphFormat.OutlineLevel = wdOutlineLevel3
                        End If

                        c = c + 1
                        d = 0
                        e = 0
                        With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "）")).Start)
                            .Delete
                            .Fields.Add Range:=.Paragraphs(1).Range, Text:="= " & c & " \* CHINESENUM3"
                            .Fields.Unlink
                            .InsertBefore Text:="（"
                        End With
' ----- 四级标题：1． -----
                    ElseIf .Text Like "1．*" Then
                        If .Text Like "1．#*" Then .Characters(1).Underline = wdUnderlineSingle
                        v5 = 0

                        If v4 = 0 Then
                            v4 = 1
                            If .End <> doc.Content.End Then
                                If .Next(4, 1) Like "1．*" Then
                                    p4 = 0
                                    If .Sentences.Count >= 2 And .Next(4, 1).Sentences.Count >= 2 Then j4 = 2 Else j4 = 1
                                Else
                                    p4 = 1
                                End If
                            End If
                        End If

                        If p4 = 1 Then
                            .Style = wdStyleHeading4
                            With .Font
                                .NameFarEast = "宋体"
                                .NameAscii = "Times New Roman"
                                .Color = wdColorGreen
                            End With
                        Else
                            .ParagraphFormat.OutlineLevel = wdOutlineLevel4
                        End If

                        d = d + 1
                        e = 0
                        doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "．")).Start).Text = d
' ----- 五级标题：（1） -----
                    ElseIf .Text Like "（1）*" Then
                        If v5 = 0 Then
                            v5 = 1
                            If .Next(4, 1) Like "（1）*" Then
                                p5 = 0
                                If .Sentences.Count >= 2 And .Next(4, 1).Sentences.Count >= 2 Then j5 = 2 Else j5 = 1
                            Else
                                p5 = 1
                            End If
                        End If

                        If p5 = 1 Then
                            .Style = wdStyleHeading5
                            With .Font
                                .NameFarEast = "宋体"
                                .NameAscii = "Times New Roman"
                                .Color = wdColorOrange
                            End With
                        Else
                            .ParagraphFormat.OutlineLevel = wdOutlineLevel5
                        End If

                        e = e + 1
                        doc.Range(Start:=.Characters(1).End, End:=.Characters(InStr(.Text, "）")).Start).Text = e
' ----- 附件段落 -----
                    ElseIf .Text Like "附件*" Or .Text Like Chr(12) & "附件*" Then
                        b = 0: c = 0: d = 0: e = 0: v4 = 0
                        GoTo sk
' ----- 阶段标题（如“第一阶段”） -----
                    ElseIf .Text Like "第" & u & "阶段*" Then
                        If v6 = 0 Then
                            v6 = 1
                            If .Next(4, 1) Like "第" & u & "阶段*" Then
                                p6 = 0
                            Else
                                p6 = 1
                            End If
                        End If
                        If .Text Like "*阶段:*" Then .Characters(InStr(.Text, ":")).Text = "："

                        If p6 = 1 Then
                            .Style = wdStyleHeading4
                            With .Font
                                .NameFarEast = "黑体"
                                .NameAscii = "Times New Roman"
                                .Color = wdColorRed
                                .Underline = wdUnderlineSingle
                            End With
                        Else
                            With doc.Range(.Start, .Start + InStr(.Text, "：")).Font
                                .Bold = True
                                .Color = wdColorRed
                                .Underline = wdUnderlineSingle
                            End With
                        End If
                        d = 0
                        GoTo sk
' ----- 序号式短语（如“第一，”） -----
                    ElseIf .Text Like "第" & u & "[， ]*" Then
                        .Characters(3).Text = "，"
                        With .Sentences(1).Font
                            .Bold = True
                            .Color = wdColorBrown
                            .Underline = wdUnderlineSingle
                        End With
                        If Not .Text Like t Then .Characters.Last.InsertBefore Text:="。"
                        GoTo sk
' ----- 一是/一要等 -----
                    ElseIf .Text Like u & "[是要]*" Then
                        If y = 0 Then
                            y = 1
                            If .Next(4, 1) Like u & "[是要]*" Then x = 1
                        End If
                        If x = 1 Then
                            With .Sentences(1).Font
                                .Bold = True
                                .Color = wdColorBrown
                                .Underline = wdUnderlineSingle
                            End With
                        End If
                        GoTo sk
' ----- 特殊字符（ASCII 31）删除 -----
                    ElseIf Asc(.Text) = 31 Then
                        .Delete
                        GoTo sk
' ----- 正文段落 -----
                    Else
                        GoTo sk
                    End If

' ----- 对于蓝色文字（即正文段落中可能出现的强调部分）进行处理 -----
                    If .Font.Color = wdColorBlue Then
                        With .Sentences(1)
                            ' 跳过电话、邮箱等含有冒号的信息
                            If .Text Like "*[电微地邮系讲][话信址箱人][：:]*" Then GoTo sk
                            If .Text Like "*[!0-9]:*" Then .Characters(InStr(.Text, ":")).Text = "："
                            If .Text Like "*[!0-9]：*" Then
                                ' 根据上下文设置不同的加粗和颜色
                                If .Previous.Text = "．" Then
                                    With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "：")).End).Font
                                        .Bold = True
                                        .Color = wdColorGreen
                                        .Underline = wdUnderlineSingle
                                    End With
                                ElseIf .Text Like "（#*" Then
                                    With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "：")).End).Font
                                        .Bold = True
                                        .Color = wdColorOrange
                                        .Underline = wdUnderlineSingle
                                    End With
                                ElseIf .Text Like "（*" Then
                                    With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "：")).End).Font
                                        .NameFarEast = "楷体"
                                        .NameAscii = "Times New Roman"
                                        .Bold = True
                                        .Color = wdColorPink
                                        .Underline = wdUnderlineSingle
                                    End With
                                ElseIf .Text Like "*" & u & "、*" Then
                                    With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "：")).End).Font
                                        .NameFarEast = "黑体"
                                        .NameAscii = "Arial"
                                        .Bold = True
                                        .Color = wdColorRed
                                        .Underline = wdUnderlineSingle
                                    End With
                                End If
                            Else
                                ' 如果段落有多句，且满足特定条件，则将首句加粗/下划线
                                If .Paragraphs(1).Range.Sentences.Count >= 2 Then
                                    If j4 = 2 Or j3 = 2 Or j2 = 2 Or j5 = 2 Then
                                        If j4 = 2 Then
                                            .Font.Color = wdColorGreen
                                        ElseIf j3 = 2 Then
                                            With .Font
                                                .Name = "楷体"
                                                .Color = wdColorPink
                                            End With
                                        ElseIf j2 = 2 Then
                                            With .Font
                                                .Name = "黑体"
                                                .Color = wdColorRed
                                            End With
                                        ElseIf j5 = 2 Then
                                            .Font.Color = wdColorOrange
                                        End If
                                        .Bold = True
                                        .Underline = wdUnderlineSingle
                                    End If
                                End If
                            End If
                        End With
' ----- 红色文字（可能是特殊标题）处理 -----
                    Else
                        .Bold = True
                        With .ParagraphFormat
                            .SpaceBefore = 13
                            .SpaceAfter = 13
                        End With
                        If .Text Like t And .Sentences.Count = 1 Then .Characters.Last.Previous.Delete

                        ' 处理冒号，将后面的文字恢复为仿宋常规蓝色
                        If .Sentences(1).Text Like "*[!0-9]:??*" Then .Characters(InStr(.Text, ":")).Text = "："
                        If .Sentences(1).Text Like "*[!0-9]：??*" Then
                            With doc.Range(Start:=.Characters(InStr(.Text, "：")).End, End:=.End).Font
                                .Name = "仿宋"
                                .Bold = False
                                .Color = wdColorBlue
                            End With
                            If .Text Like "*[．：]*" Then doc.Range(.Start + InStr(.Text, "．"), .Start + InStr(.Text, "：")).Underline = wdUnderlineSingle
                        Else
                            If .Sentences.Count >= 2 Then
                                With doc.Range(Start:=.Sentences(1).End, End:=.End).Font
                                    .Name = "仿宋"
                                    .Bold = False
                                    .Color = wdColorBlue
                                End With
                                .Sentences(1).Underline = wdUnderlineSingle
                                With .ParagraphFormat
                                    .SpaceBefore = 0
                                    .SpaceAfter = 0
                                End With
                            End If
                        End If
                    End If
                End With
sk:
            Next

RangeFormat:
            ' 对处理后的区域设置统一的段落格式和字体
            If .Start <> 0 Then
                .InsertParagraphBefore
                .Paragraphs(1).Range.Style = wdStyleNormal
                z = 1
            End If

            With .Font
                .Size = 16
                .Kerning = 0
                .DisableCharacterSpaceGrid = True
            End With

            With .ParagraphFormat
                .LineSpacing = LinesToPoints(1.5)
                .Alignment = wdAlignParagraphJustify
                .CharacterUnitFirstLineIndent = 2
                .KeepWithNext = False
                .KeepTogether = False
                .AutoAdjustRightIndent = False
                .DisableLineHeightGrid = True
            End With

            If z = 1 Then
                With .Paragraphs(1).Range
                    If .Next(4, 1).ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText Then .Font.Size = 6 Else .Font.Size = 9
                End With
            End If

            If y = 1 Then
                With .ParagraphFormat
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                    .LineSpacing = LinesToPoints(1.25)
                End With
            End If
        End With
    Next
End Function

' 设置正文样式（宋体10.5pt，两端对齐，无首行缩进等）
Function Word2003NormalStyle()
    With ActiveDocument.Styles("正文")
        With .Font
			.NameFarEast = "宋体"                      ' 设置东亚文字（中文、日文、韩文）字体为宋体
			.NameAscii = "Times New Roman"             ' 设置 ASCII 字符（英文、数字等）字体为 Times New Roman
			.NameOther = "Times New Roman"             ' 设置其他字符（如符号、非 ASCII 拉丁字母）字体为 Times New Roman
			.Name = "宋体"                             ' 同时设置所有字体的默认名称（优先覆盖以上三项）
			.Size = 10.5                               ' 字体大小 10.5 磅
			.Bold = False                              ' 取消加粗
			.Italic = False                            ' 取消斜体
			.Underline = wdUnderlineNone               ' 无下划线
			.UnderlineColor = wdColorAutomatic         ' 下划线颜色自动（跟随文字颜色）
			.StrikeThrough = False                     ' 无单删除线
			.DoubleStrikeThrough = False               ' 无双删除线
			.Outline = False                           ' 无镂空轮廓效果
			.Emboss = False                            ' 无阳文浮雕效果
			.Shadow = False                            ' 无阴影
			.Hidden = False                            ' 不隐藏文字
			.SmallCaps = False                         ' 不转换为小型大写字母
			.AllCaps = False                           ' 不转换为全部大写字母
			.Color = wdColorAutomatic                  ' 字体颜色自动（通常为黑色）
			.Engrave = False                           ' 无阴文雕刻效果
			.Superscript = False                       ' 无上标
			.Subscript = False                         ' 无下标
			.Scaling = 100                             ' 字符缩放比例 100%（正常宽度）
			.Kerning = 1                               ' 启用字距调整的最小字号为 1 磅（实际表示始终调整）
			.Animation = wdAnimationNone               ' 无文字动画效果
			.DisableCharacterSpaceGrid = False         ' 不禁用字符网格对齐（保留网格对齐）
			.EmphasisMark = wdEmphasisMarkNone         ' 无强调标记（如着重号）
        End With

        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = wdBaselineAlignAuto
        End With

        .NoSpaceBetweenParagraphsOfSameStyle = False
        .ParagraphFormat.TabStops.ClearAll
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "正文"
    End With
End Function

' 设置“称呼”样式（楷体、加粗、粉色、下划线、无缩进）
Function HelloStyle()
    With Selection
        With .Font
            .Name = "楷体"
            .Bold = True
            .Color = wdColorPink
            .Underline = wdUnderlineSingle
        End With
        With .ParagraphFormat
            .CharacterUnitFirstLineIndent = 0
            .FirstLineIndent = CentimetersToPoints(0)
        End With
    End With
End Function

' 处理一级标题（主标题），设置样式、字体、居中、字符间距等
Function Title1()
    With ActiveDocument.Paragraphs(1).Range
        ' 将主标题扩展到包含整个标题段落（直到遇到标点或下一级标题）
        Do While Not (.Next(4, 1) Like "*[。：；，！？…—.:;,!?]*" Or .Next(4, 1) Like "[一1][、.．]*" Or .Next(4, 1) Like "第[一1]*" Or .Next(4, 1) Like "（[一1]）*" Or .Next.Information(12) Or .Next(4, 1) Like "序*言?")
            .MoveEnd 4
        Loop
        .InsertParagraphBefore
        .Characters.Last.InsertParagraphBefore
        .Style = wdStyleHeading1
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone   ' 取消下边框
        With .Font
            .Name = "方正小标宋简体"
            .Size = 22
            .Bold = False
            .Color = wdColorAutomatic
            .Kerning = 0
        End With
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .Space1
            .Alignment = wdAlignParagraphCenter
            .AutoAdjustRightIndent = False
            .DisableLineHeightGrid = True
        End With
        With .Characters
            .First.Font.Size = 21   ' 首字符略小
            .Last.Font.Size = 26    ' 末字符略大（视觉调整）
        End With

        ' 处理主标题后面的称呼部分
        With .Next(4, 1)
            If .Information(12) Then  ' 如果是表格
                .Previous.Delete
                .Parent.Characters.First.Delete
                .Previous(4, 1).ParagraphFormat.LineUnitAfter = 1
                .Previous(4, 1).Select
                ExtTitle   ' 调整标题间距
                Exit Function
            Else
                If .Text Like "*[：:]?" Then
                    .Characters.Last.Previous.Text = "："
                    If .ComputeStatistics(wdStatisticLines) <= 2 Then
                        .Select
                        HelloStyle   ' 应用称呼样式
                    End If
                ElseIf .Next.Information(12) Then
                    With .ParagraphFormat
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = CentimetersToPoints(0)
                        .LineSpacing = LinesToPoints(1.25)
                    End With
                    .Previous.Font.Size = 16
                    .Parent.Paragraphs(1).Range.Delete
                    Exit Function
                End If
            End If
        End With

        ' 删除多余空格，统一括号格式
        If .Next(4, 1) Like "*：?" Then .MoveEnd 4
        With .Find
            .Execute "(", , , 0, , , , , , "（", 2
            .Execute ")", , , 0, , , , , , "）", 2
            .Execute "[ 　^s^t]", , , 1, , , , , , "", 2
        End With
        If .Text Like "*：?" Then .MoveEnd 4, -1

        ' 处理副标题（讨论稿、日期等）
        With .Paragraphs.Last.Range.Previous(4, 1)
            If .Text Like "（*）?" Then
                .Select
                SubTitle
                If Len(.Text) = 5 Then
                    .Characters(2).InsertAfter Text:=" "
                ElseIf Len(.Text) = 6 Then
                    .Characters(2).InsertAfter Text:=" "
                    .Characters(4).InsertAfter Text:=" "
                End If
                With .Next(4, 1)
                    If .Next(4, 1).Text Like "一、*" Then .Font.Size = 16
                End With

            ElseIf .Text Like "*年*月*日*" Then
                .Select
                SubTitle

            ElseIf .Text Like "*的???" Then
                If .ComputeStatistics(wdStatisticLines) >= 2 Then
                    .Characters.Last.Select
                    With Selection
                        .HomeKey 5, 1
                        If Len(.Text) = 3 Then
                            .Characters(InStrRev(.Text, "的")).InsertParagraphAfter
                            .MoveEnd 4, -1
                            .Expand 4
                            CutLine
                            .Next(4, 1).Select
                            ExtTitle
                        ElseIf Len(.Text) = 1 Then
                            .Expand 4
                            .Paragraphs(1).Range.Characters(InStrRev(.Text, "的")).InsertParagraphAfter
                            .MoveStart 4
                            ExtTitle
                        End If
                    End With
                End If
            Else
                .Select
                ExtTitle
            End If

            With .Previous(4, 1)
                If Len(.Text) <> 1 Then
                    .Characters.Last.Select
                    With Selection
                        .HomeKey 5, 1
                        If Len(.Text) = 1 Then
                            .Expand 4
                            CutLine
                        End If
                    End With
                End If
            End With
        End With

        ' 处理引号等标点导致的悬挂缩进
        Dim i As Paragraph
        For Each i In .Paragraphs
            With i.Range
                If .Text Like "[“（《〈]*[!）]?" Then
                    .ParagraphFormat.CharacterUnitLeftIndent = -0.5
                ElseIf .Text Like "[!（]*[）”〉》]?" Then
                    .InsertBefore Text:=" "
                End If
            End With
        Next
    End With
End Function

' 调整标题间距（在文字之间插入空格，使标题宽度适中）
Function ExtTitle()
    With Selection
        If .Type = wdSelectionIP Then
            .HomeKey
            .EndKey 5, 1
        End If
        If Len(.Text) = 3 Then
            .Characters(1).InsertAfter Text:="    "
        ElseIf Len(.Text) = 4 Then
            .Characters(1).InsertAfter Text:="   "
            .Characters(5).InsertAfter Text:="   "
            If .Text Like "协*议*书?" Then .Font.Size = 26
        ElseIf Len(.Text) = 5 Then
            .Characters(1).InsertAfter Text:="  "
            .Characters(4).InsertAfter Text:="  "
            .Characters(7).InsertAfter Text:="  "
        ElseIf Len(.Text) = 6 Then
            .Characters(1).InsertAfter Text:="  "
            .Characters(4).InsertAfter Text:="  "
            .Characters(7).InsertAfter Text:="  "
            .Characters(10).InsertAfter Text:="  "
        ElseIf Len(.Text) = 7 Then
            .Characters(1).InsertAfter Text:=" " & ChrW(160)
            .Characters(4).InsertAfter Text:=" " & ChrW(160)
            .Characters(7).InsertAfter Text:=" " & ChrW(160)
            .Characters(10).InsertAfter Text:=" " & ChrW(160)
            .Characters(13).InsertAfter Text:=" " & ChrW(160)
        End If
    End With
End Function

' 设置副标题样式（楷体、18pt、蓝绿色等）
Function SubTitle()
    With Selection
        .Expand 4
        .Style = wdStyleSubtitle
        With .Font
            .NameFarEast = "楷体"
            .NameAscii = "Times New Roman"
            .Size = 18
            .Color = wdColorTeal
            .Kerning = 0
        End With
        With .ParagraphFormat
            .AutoAdjustRightIndent = False
            .DisableLineHeightGrid = True
        End With
    End With
End Function

' 处理落款（日期、单位、附件等）
Function Inscribe()
    Const FindDate As String = "^13[2２二][0-9０-９ 　^s^t一二三四五六七八九十零〇○Oo0０Ｏｏ]@[!^13]@[日月0-9０-９][^13^12]"
    Const u As String = "[一二三四五六七八九十]"
    Dim doc As Document, arr, brr, crr, i As Paragraph, n&, t$, x&, y!, w&
    Dim NumDate&, oDate1 As Range, oDate2 As Range, rngDate As Range

    Set doc = ActiveDocument

    ' 查找日期段落（如“2023年1月1日”或“二〇二三年一月一日”）
    With doc.Content.Find
        Do While .Execute(FindDate, , , 1, , , 1)
            With .Parent
                If Len(.Text) = 0 Then Exit Do
                .MoveStart
                If .Previous(4, 1).ParagraphFormat.OutlineLevel <> wdOutlineLevelBodyText Then GoTo sk
                If .Text Like "*" & Chr(12) Then .MoveEnd 1, -1
                t = .Text
                t = Replace(t, " ", "")
                t = Replace(t, "　", "")
                t = Replace(t, vbTab, "")
                t = Replace(t, ChrW(160), "")
                If .End = doc.Content.End Then t = Replace(t, vbCr, "")
                .Text = t

                ' 将中文日期转换为阿拉伯数字日期
                If .Text Like "####年*月*" And Len(.Text) <= 12 Then
                    .CharacterWidth = wdWidthHalfWidth
                ElseIf .Text Like "####[!0-9０-９]*#*" And Len(.Text) <= 11 Then
                    .CharacterWidth = wdWidthHalfWidth
                    .Characters(5).Text = "年"
                    If .Text Like "*年#?" Or .Text Like "*年##?" Then
                        .Characters.Last.InsertBefore Text:="月"
                    Else
                        If .Text Like "*年#?#?" Or .Text Like "*年#?##?" Then
                            .Characters(7).Text = "月"
                        ElseIf .Text Like "*年##?#?" Or .Text Like "*年##?##?" Then
                            .Characters(8).Text = "月"
                        End If
                        .Characters.Last.InsertBefore Text:="日"
                    End If
                ElseIf .Text Like "二???年*月*" And Len(.Text) <= 13 Then
                    brr = Split("一/二/三/四/五/六/七/八/九/零/〇/○/O/o/0/０/Ｏ/ｏ", "/")
                    crr = Split("1/2/3/4/5/6/7/8/9/0/0/0/0/0/0/0/0/0", "/")
                    t = .Text
                    For n = 0 To 17
                        t = Replace(t, brr(n), crr(n))
                    Next
                    If t Like "*年十月*" Then
                        t = Replace(t, "十", "10", 1, 1)
                    ElseIf t Like "*年十?月*" Then
                        t = Replace(t, "十", "1", 1, 1)
                    End If
                    If t Like "*月?十?日?" Then
                        t = Replace(t, "十", "")
                    ElseIf t Like "*月十?日?" Then
                        t = Replace(t, "十", "1")
                    ElseIf t Like "*月?十日?" Then
                        t = Replace(t, "十", "0")
                    End If
                    If .End = doc.Content.End Then t = Replace(t, vbCr, "")
                    .Text = t
                    n = 0
                Else
                    GoTo sk
                End If

                ' 补全月和日中的个位数（如“1月1日”改为“01月01日”？实际上代码是补全“1日”为“1日”？看代码逻辑）
                If .Text Like "*月?" Then .Characters.Last.InsertBefore Text:="1日"
                If .Text Like "*0?月*" Then .Characters(6).Delete
                If .Text Like "*##月0?日*" Then .Characters(9).Delete
                If .Text Like "*#月0?日*" Then .Characters(8).Delete
                If .Font.Size = 22 Or .Font.Size = 18 Then GoTo sk

                ' 设置日期格式：粉色、加粗、右对齐
                .Font.Color = wdColorPink
                .Bold = True
                With .ParagraphFormat
                    .Alignment = wdAlignParagraphRight
                    .CharacterUnitRightIndent = 5.9
                End With
                t = .Text

                ' 处理日期前面的单位名称
                With .Previous(4, 1)
                    If .Text Like "*[!。：；，、！？…—.:;,!?]?" And .Text Like "[一-﨩]*" Then
                        .Font.Color = wdColorRed
                        .Bold = True
                        x = Len(.Text) - 1
                        If x >= 17 Then x = 16
                        arr = Array(17.43, 16.93, 16.4, 15.92, 15.44, 14.96, 14.26, 13.65, 13.12, 12.55, 12.07, 11.6, 11.06, 10.5, 9.56)
                        y = arr(x - 2)
                        If t Like "*年#月##日?" Or t Like "*年##月#日?" Then
                            y = y - 0.15
                        ElseIf t Like "*年##月##日?" Then
                            y = y - 0.4
                        End If
                        .InsertBefore Text:=vbCr & vbCr & vbCr
                        .MoveStart 1, 3
                        .ParagraphFormat.CharacterUnitFirstLineIndent = y

                        NumDate = NumDate + 1
                        If NumDate = 1 Then
                            Set oDate1 = doc.Range(.Start, .End).Next(4, 1)
                        ElseIf NumDate = 2 Then
                            Set oDate2 = doc.Range(.Start, .End).Next(4, 1)
                        End If
                    Else
                        .InsertParagraphAfter
                    End If
                End With
sk:
                .Start = .End
            End With
        Loop
    End With

    ' 根据找到的日期数量，处理落款和附件
    If NumDate = 1 Then
        If oDate1.End = doc.Content.End Then
            If Not doc.Content Like "*" & vbCr & "附件*" Then Exit Function
        End If
        Set rngDate = oDate1

    ElseIf NumDate = 0 Then
        Exit Function

    ElseIf NumDate >= 2 Then
        ' 如果有两个日期，则复制第一个日期的格式到第二个日期，并重新编号二级标题等
        Set rngDate = oDate2
        doc.Paragraphs(2).Range.Select
        Selection.CopyFormat
        With oDate1.Next(4, 1)
            Do While Not (.Next(4, 1) Like "*[。：；，！？….:;,!? 　^s^t]*" Or .Next(4, 1) Like "[一1][、.．]*" Or .Next(4, 1) Like "（[一1]）*" Or .Next(4, 1) Like "第[一1]*" Or .Next.Information(12))
                .MoveEnd 4
            Loop
            .InsertParagraphBefore
            .Characters.Last.InsertParagraphBefore
            .Select
            With Selection
                .PasteFormat
                .Characters.First.Font.Size = 21
                .Characters.Last.Font.Size = 26
                .InsertBreak Type:=2
            End With
            With .Next(4, 1)
                If .Text Like "*[：:]?" Then
                    .Characters.Last.Previous.Text = "："
                    If .ComputeStatistics(wdStatisticLines) <= 2 Then
                        .Select
                        HelloStyle
                    End If
                End If
            End With
        End With

        ' 重新对二级标题进行自动编号
        For Each i In doc.Range(doc.Sections(2).Range.Start, rngDate.Start).Paragraphs
            With i.Range
                If .Text Like u & "、*" Or .Text Like u & u & "、*" Then
                    w = w + 1
                    With doc.Range(Start:=.Start, End:=.Characters(InStr(.Text, "、")).Start)
                        .Delete
                        .Fields.Add Range:=.Paragraphs(1).Range, Text:="= " & w & " \* CHINESENUM3"
                        .Fields.Unlink
                    End With
                End If
            End With
        Next

        With rngDate
            If Not .Text Like "*日" & vbCr Then
                .Characters(InStr(.Text, "日")).InsertParagraphAfter
                If .Paragraphs.Count = 2 Then .MoveEnd 4, -1
            End If
        End With

        ' 删除第一个日期之前的附件
        With doc.Range(0, oDate1.End).Find
            Do While .Execute("^13附件*^13", , , 1, , , 1)
                With .Parent
                    .MoveStart
                    Do While .Next(4, 1) <> vbCr
                        .MoveEnd 4
                    Loop
                    .Delete
                    Exit Do
                End With
            Loop
        End With

        ' 调整引子部分的段落间距
        With doc.Range(0, oDate1.End)
            Do
                .MoveStart 4
            Loop Until .Previous(4, 1).Font.Size = 26
            .MoveEnd 4, -5
            If .ComputeStatistics(wdStatisticLines) <= 4 Then
                .ParagraphFormat.LineSpacing = LinesToPoints(1.75)
                .Previous.Font.Size = 28
            End If
        End With
    End If

    ' 处理附件部分
    Dim rngBefore As Range, rngAfter As Range, Before&, After&, j&, s$, Check&, v&

'Before: 查找日期之前的附件
    Set rngBefore = doc.Range(Start:=0, End:=rngDate.Start)
    With rngBefore.Find
        Do While .Execute("^13附件*^13", , , 1, , , 1)
            With .Parent
                .MoveStart
                Do While .Next(4, 1) <> vbCr
                    .MoveEnd 4
                Loop
                .InsertParagraphBefore
                .MoveStart
                Before = 1
                Exit Do
            End With
        Loop
    End With
    If rngDate.End = doc.Content.End Then GoTo ds

After:
    ' 处理日期之后的附件
    Set rngAfter = doc.Range(Start:=rngDate.End - 1, End:=doc.Content.End)
    With rngAfter
        With .Find
            .Execute "^12^13", , , 0, , , , , , "^12", 2
            .Execute "^12^12", , , 0, , , , , , "^12", 2
        End With
        If .Paragraphs.Count <= 5 Then
            Dim NoAttach&
            .MoveStart
            s = .Text
            .Delete
            After = 1
            NoAttach = 1
            GoTo ds
        End If
    End With

    With rngAfter.Find
        Do While .Execute("[^13^12]附件*^13", , , 1, , , 1)
            With .Parent
                .MoveStart

                ' 处理附件标题中的冒号
                If .Text Like "*:??*" Then .Characters(InStr(.Text, ":")).Text = "："
                If .Text Like "*：??*" Then
                    .Characters(InStr(.Text, "：")).InsertParagraphAfter
                    .MoveEnd 4, -1
                End If

                ' 处理分页符等
                If Asc(.Previous) = 13 Then
                    .InsertBreak Type:=0
                    .MoveEnd 4
                    .MoveEnd 1, -1
                ElseIf Asc(.Previous) = 12 Then
                    .MoveEnd 1, -1
                End If
                If .Previous.Previous.Previous.Information(12) Then .Previous.Previous.Delete

                n = n + 1
                .Text = "附件" & n
                If Asc(.Previous) = 13 Then .Previous.Delete

                ' 设置附件标号样式
                With .Font
                    .NameFarEast = "黑体"
                    .NameAscii = "Times New Roman"
                    .Size = 16
                    .Bold = True
                    .Color = wdColorRed
                End With
                With .ParagraphFormat
                    .CharacterUnitFirstLineIndent = 0
                    .FirstLineIndent = CentimetersToPoints(0)
                    .Alignment = wdAlignParagraphJustify
                    .Space1
                End With

                ' 处理附件标题正文
                With .Next(4, 1)
                    If .Information(12) Then
                        .MoveEnd 1, -1
                        .Previous.InsertBefore Text:=vbCr & .Text
                        .Select
                        Selection.Rows(1).Delete
                        .MoveStart 4, -1
                    Else
                        Do While Not (.Next(4, 1) Like "*[`。：；，！？….:;,!?]*" Or .Next(4, 1) Like "[一1][、.．]*" Or .Next(4, 1) Like "（[一1]）*" Or .Next(4, 1) Like "第[一1]*" Or .Next.Information(12))
                            .MoveEnd 4
                        Loop
                        With .Find
                            .Execute "(", , , 0, , , , , , "（", 2
                            .Execute ")", , , 0, , , , , , "）", 2
                            .Execute "[ 　^s^t]", , , 1, , , , , , "", 2
                        End With
                    End If
                    With .Font
                        .Name = "方正小标宋简体"
                        .Size = 20
                        .Bold = False
                        .Color = wdColorAutomatic
                    End With
                    With .ParagraphFormat
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = CentimetersToPoints(0)
                        .Alignment = wdAlignParagraphCenter
                        .Space1
                    End With
                    t = Replace(.Text, vbCr, "")
                    t = t & vbCr
                    s = s & t
                    If .Sections(1).PageSetup.Orientation = wdOrientPortrait Then
                        .InsertParagraphBefore
                        .Characters.Last.InsertParagraphBefore
                        If Not .Next.Information(12) Then .Paragraphs.Last.Range.ParagraphFormat.LineUnitAfter = 0.5
                        .MoveStart
                        .MoveEnd 1, -1
                    Else
                        With .ParagraphFormat
                            .LineUnitBefore = 0.5
                            .LineUnitAfter = 1
                        End With
                    End If
                    If .Text Like "[“（《〈]*[!）]?" Then
                        .ParagraphFormat.CharacterUnitLeftIndent = -0.5
                    ElseIf .Text Like "*[）”〉》]?" Then
                        .InsertBefore Text:=" "
                    End If
                End With
                After = 1
                .Start = .End
            End With
        Loop
    End With
    If n = 1 Then rngAfter.Previous.Delete

' 如果没找到附件标号，尝试在日期后直接添加“附件”
    If After = 0 Then
        With rngDate.Next(4, 1)
            If .Text = Chr(12) Then
                .InsertAfter Text:="附件" & vbCr
                GoTo After

            ElseIf Asc(.Previous) = 13 Then
                .InsertBefore Text:="附件" & vbCr
                GoTo After

            ElseIf Asc(.Previous) = 12 Then
                .Select
                If .Information(12) Then
                    With Selection
                        .SplitTable
                        .TypeText Text:="附件"
                    End With
                    GoTo After
                Else
                    Stop
                End If
            End If
        End With
    End If

' 整理附件列表的格式
ds:
    With rngBefore
        If Before = 1 Then
            If After = 1 Then
                If .Text Like "附件[：:]" & vbCr & "*" Then .Paragraphs(1).Range.Delete
                If .Paragraphs.Count <> n Then Check = 1
                .Text = s
            End If
        Else
            If After = 1 Then
                With rngDate
                    .MoveStart 4, -4
                    .InsertBefore Text:=vbCr & s
                    .MoveStart
                    .MoveEnd 4, -5
                End With
                .SetRange rngDate.Start, rngDate.End
            Else
                If Not rngDate.Next.Text = "附" Then
                    rngDate.InsertAfter Text:="附件" & vbCr
                End If
            End If
        End If
        If .Previous.Previous(4, 1).ParagraphFormat.SpaceAfter = 13 Then .Previous.Font.Size = 9
        With .Font
            .Bold = False
            .Color = wdColorViolet
            .Underline = wdUnderlineNone
        End With

        ' 简化附件列表中的多余字符
        For Each i In .Paragraphs
            With i.Range
                If .Text Like "附*" Then .Characters(1).Delete
                If .Text Like "件*" Then .Characters(1).Delete
                If .Text Like "[：:]*" Then .Characters(1).Delete
                If .Text Like "#[、.．]*" Then .Characters(1).Delete
                If .Text Like "[、.．]*" Then .Characters(1).Delete
            End With
        Next

        ' 设置附件列表的缩进和编号
        If n = 0 Then
            If rngBefore.Paragraphs.Count >= 2 Then n = 2 Else n = 1
        End If

        If n = 1 Then
            .InsertBefore Text:=vbTab
        Else
            For Each i In .Paragraphs
                j = j + 1
                i.Range.InsertBefore Text:=j & "．" & vbTab
            Next
        End If
        With .ParagraphFormat
            .CharacterUnitLeftIndent = 7.65
            .CharacterUnitFirstLineIndent = -1.56
        End With
        .InsertBefore Text:="附件："
        With .Paragraphs(1).Range.ParagraphFormat
            .CharacterUnitLeftIndent = 3.05
            .CharacterUnitFirstLineIndent = -4.62
        End With
        If n = 1 Then .ParagraphFormat.CharacterUnitFirstLineIndent = -3.1

        ' 设置“附件：”字样样式
        With doc.Range(.Start, .Start + 3).Font
            .Bold = True
            .Color = wdColorRed
        End With

        ' 微调字符间距，使附件列表美观
        If n = 1 Then
            If .ComputeStatistics(wdStatisticLines) <> 1 Then
                .Characters.Last.Select
                With Selection
                    .HomeKey 5, 1
                    If Len(.Text) = 1 Then
                        .Expand 4
                        .MoveStart 1, 5
                        If .Text Like "0##*" Then .Font.Spacing = -0.4 Else .Font.Spacing = -0.2
                    End If
                End With
            End If
        Else
            For Each i In .Paragraphs
                With i.Range
                    If .ComputeStatistics(wdStatisticLines) <> 1 Then
                        .Characters.Last.Select
                        With Selection
                            .HomeKey 5, 1
                            If Len(.Text) = 1 Then
                                .Expand 4
                                If .Text Like "附*" Then .MoveStart 1, 7 Else .MoveStart 1, 4
                                .Font.Spacing = -0.3
                            End If
                        End With
                    End If
                End With
            Next
        End If
    End With

    If Check = 1 Then MsgBox "请检查附件！", 0 + 16

    BackUnderLine   ' 恢复下划线（将“`”还原为带下划线的空格）
    CheckTables     ' 处理表格相关格式
End Function

' 将之前替换的“`”还原为带下划线的空格
Function BackUnderLine()
    With ActiveDocument.Content.Find
        .Replacement.Font.Underline = wdUnderlineSingle
        .Execute FindText:="`", ReplaceWith:=" ", Format:=True, Replace:=wdReplaceAll
    End With
End Function

' 检查并调整表格前后的格式
Function CheckTables()
    Dim doc As Document, t As Table, z&
    Set doc = ActiveDocument
    For Each t In doc.Tables
        With t.Range
            With .Previous(4, 1)
                If .Font.Size = 20 Then
                    z = 1 ' 附件表格
                ElseIf .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText Then
                    If .Previous.Font.Size = 20 Then
                        With .ParagraphFormat
                            .CharacterUnitFirstLineIndent = 0
                            .FirstLineIndent = CentimetersToPoints(0)
                            .LineSpacing = LinesToPoints(1.25)
                        End With
                        z = 1
                    End If
                End If
            End With
            With .Next(4, 1)
                If .Font.Size = 1 Then Exit Function
                If z = 1 Then
                    If .Text Like Chr(12) & "附件*" Then
                    ElseIf .End <> doc.Content.End Then
                        If .Next.ParagraphFormat.OutlineLevel <> wdOutlineLevelBodyText Then GoTo pass
                        If .Next = Chr(12) Then .Delete
                        If .MoveEndUntil(Chr(12)) = 0 Then .MoveEndUntil (vbCr)
                        With .ParagraphFormat
                            .CharacterUnitFirstLineIndent = 0
                            .FirstLineIndent = CentimetersToPoints(0)
                            .LineSpacing = LinesToPoints(1.25)
                        End With
                        .MoveStart
                        .ParagraphFormat.CharacterUnitFirstLineIndent = -1.96
                    End If
                End If
            End With
        End With
pass:
    Next
End Function

' 设置页码（公文格式：一字线居中，页码数字）
Function PageNumber()
    Dim s As Section, i&

    With ActiveDocument
        If .ComputeStatistics(wdStatisticPages) >= 3 Then i = 1

        For Each s In .Sections
            With s.Footers(wdHeaderFooterPrimary)
                .Range.Delete
                .LinkToPrevious = True

                If i = 1 Then
                    With .PageNumbers
                        .NumberStyle = wdPageNumberStyleArabic
                        .RestartNumberingAtSection = False
                        .Add PageNumberAlignment:=wdAlignPageNumberCenter
                    End With
                End If
            End With
        Next

        If i = 1 Then
            With .Sections(1).Footers(1).Range
                .InsertBefore Text:=ChrW(8212) & " "   ' 左一字线
                .Characters.Last.Previous.InsertAfter Text:=" " & ChrW(8212) ' 右一字线
                With .Font
                    .Name = "宋体"
                    .Size = 14
                End With
            End With
        End If

        With .Sections(1)
            .Headers(1).Range.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            With .PageSetup
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
            End With
        End With
    End With
End Function

' 根据文档页数设置窗口显示方式（多页显示或单页）
Function NumPages()
    Dim p&
    p = ActiveDocument.ComputeStatistics(wdStatisticPages)
    With ActiveWindow.ActivePane.View.Zoom
        If p < 99 Then
            If .PageColumns = p Then .PageColumns = 3 Else .PageColumns = p
        Else
            If .PageColumns = 15 Then .PageColumns = 3 Else .PageColumns = 15
        End If
        .PageRows = 1
    End With
    If p = 1 Then ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
End Function

' 缩减一行（通过减小字符间距）
Function CutLine()
    Dim r As Range, s As Range, i&, j&
    With Selection
        If .Information(12) Then
            .MoveEndUntil cset:=vbCr, Count:=wdForward
            .MoveStart 5, -1
            Set r = .Range
            Set s = .Range
        Else
            Set r = .Paragraphs(1).Range
            Set s = .Paragraphs(1).Range
        End If
        j = .Paragraphs(1).Range.ParagraphFormat.Alignment
    End With
    With s
        With .Font
            .Spacing = 0
            .Scaling = 100
        End With
        i = .ComputeStatistics(1)
        If i = 1 Then Exit Function
    End With
    With r.Font
        If j <> 1 Then .Parent.MoveStart
        Do
            .Spacing = .Spacing - 0.1
        Loop Until s.ComputeStatistics(1) = i - 1
    End With
End Function

' 取消字符网格对齐
Function CancelGrid()
    With Selection
        If .Type = wdSelectionIP Then SelectText
        With .Font
            .Kerning = 0
            .DisableCharacterSpaceGrid = True
        End With
        With .ParagraphFormat
            .AutoAdjustRightIndent = False
            .DisableLineHeightGrid = True
        End With
    End With
End Function

' 选择正文（从文号之后到附件之前）—— 虽然未被主过程调用，但被 CancelGrid 调用，故保留
Function SelectText()
    With ActiveDocument.Content
        If .Paragraphs(2).Range Like "*〔????〕*号?" Then .MoveStart 4, 2
        Do
            .MoveStart 4
        Loop Until .Text Like vbCr & "*"
        .MoveStart 4
        .Select
        With .Find
            .Execute "^12附件", , , 1, , , 1
            If .Found = True Then
                .Parent.Parent.Range(Start:=Selection.Start, End:=.Parent.End - 2).Select
            End If
        End With
    End With
End Function