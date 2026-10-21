Sub FormatOfficialSignature()
    ' 功能：将选中的文本或光标所在段落设置为公文落款格式
    ' 要求：请将发文机关署名和日期放在连续的两个段落中，并同时选中它们，或将光标置于署名段落。
    Dim signaturePara As Paragraph, datePara As Paragraph
    Dim signatureRange As Range, dateRange As Range
    Dim signatureLen As Integer, dateLen As Integer
    Dim targetParagraphs As Paragraphs

    ' --- 步骤1：智能选择要操作的段落 ---
    ' 如果用户有选中区域，则操作选中的段落；否则操作当前光标所在的段落及其下一段（假设是落款的两行）
    If Selection.Paragraphs.Count >= 2 Then
        Set targetParagraphs = Selection.Paragraphs
    Else
        ' 如果只选中或光标在一段，默认操作当前段和下一段
        Set signaturePara = Selection.Paragraphs(1)
        ' 确保不超出文档范围
        If signaturePara.Range.Next Is Nothing Then
            MsgBox "请同时选中署名和日期两行，或将光标置于署名行。", vbExclamation, "提示"
            Exit Sub
        End If
        Set datePara = signaturePara.Next
        ' 重新构造一个包含两个段落的集合用于处理
        Set signaturePara = Nothing ' 后面会重新赋值
        GoTo ProcessParas
    End If

    ' 确保我们有两个段落
    If targetParagraphs.Count < 2 Then
        MsgBox "请至少选中包含署名和日期的两个段落。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 将前两个段落分别指定为署名和日期
    Set signaturePara = targetParagraphs(1)
    Set datePara = targetParagraphs(2)

ProcessParas:
    ' --- 步骤2：获取段落文本并计算长度 ---
    Set signatureRange = signaturePara.Range
    Set dateRange = datePara.Range

    ' 去掉段落标记（回车符）以获取纯文本长度
    signatureLen = Len(Trim(signatureRange.Text))
    dateLen = Len(Trim(dateRange.Text))

    ' 检查是否获取到内容
    If signatureLen = 0 Or dateLen = 0 Then
        MsgBox "未检测到署名或日期内容。", vbExclamation, "提示"
        Exit Sub
    End If

    ' --- 步骤3：设置段落格式（核心步骤）---
    With datePara.Range.ParagraphFormat
        .Alignment = wdAlignParagraphRight      ' 日期右对齐
        .RightIndent = CentimetersToPoints(2)   ' ★ 右缩进4个字符（通常设为2厘米左右，可根据字体微调）
    End With

    With signaturePara.Range.ParagraphFormat
        .Alignment = wdAlignParagraphRight      ' 署名也右对齐
        ' ★ 关键：根据署名长度动态调整右缩进，使其大致居中于日期
        ' 公式：让署名在日期上方居中对齐 ≈ 日期的右缩进 + (日期文本宽度 - 署名文本宽度)/2
        ' 这里用简化的估算：字数差 * 一个字宽（约0.75厘米），可根据实际效果调整系数
        Dim indentAdjust As Single
        indentAdjust = CentimetersToPoints((dateLen - signatureLen) * 0.75 / 2)
        .RightIndent = CentimetersToPoints(2) + indentAdjust
    End With

    MsgBox "落款格式设置完成！", vbInformation, "成功"

End Sub