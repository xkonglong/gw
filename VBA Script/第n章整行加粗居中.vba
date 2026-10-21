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

Sub 第n章整行加粗居中()
    Dim para As Paragraph
    Dim txt As String
    Dim pos As Integer

    ' 遍历当前文档所有段落
    For Each para In ActiveDocument.Paragraphs
        txt = para.Range.Text
        If Len(txt) > 0 Then
            txt = Left(txt, Len(txt) - 1)   ' 去掉段落标记
        End If
        
        ' 判断是否以“第”开头，且前几个字符内包含“章”
        If Left(txt, 1) = "第" Then
            pos = InStr(txt, "章")
            If pos > 1 And pos <= 5 Then     ' 限制“章”出现在前5个字符内
				' 3. 顶格（清除所有缩进）
				With para.Range.ParagraphFormat
                    .LeftIndent = 0              ' 左缩进归零
                    .RightIndent = 0             ' 右缩进归零
                    .FirstLineIndent = 0         ' 首行缩进归零（即“顶格”）
                    .CharacterUnitLeftIndent = 0 ' 字符单位左缩进归零（兼容）
                    .CharacterUnitFirstLineIndent = 0
                End With
                ' 1. 居中
                para.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                ' 2. 加粗
                para.Range.Font.Bold = True
            End If
        End If
    Next para
    
    MsgBox "处理完成！所有“第…章”标题已设置为：居中、加粗、顶格（无缩进）。", vbInformation
End Sub