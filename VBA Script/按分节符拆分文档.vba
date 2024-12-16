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
'  作者: 小恐龙
'  日期: 2024年1月31日
' ver1.1 修复msgbox提示出错的bug


Sub xkonglong_根据分节符拆分文档()

    '本段代码作用: 如果没有分节符,提示后退出
    If Word.ActiveDocument.Sections.Count <= 1 Then
        MsgBox "拆分文档前请插入分节符!"
        Exit Sub
    End If


    '本段代码作用: 选择拆分后要保存的文件夹'
    Dim dia As FileDialog
    Set dia = Application.FileDialog(msoFileDialogFolderPicker)
    dia.Title = "请选择拆分后要保存的文件夹"
    If Word.ActiveDocument.Saved Then dia.InitialFileName = Word.ActiveDocument.path
    If dia.Show = 0 Then Exit Sub
    Dim sPath As String
    sPath = dia.SelectedItems(1)
    
    
    '本段代码作用:拆分文档并保存
    Dim oDoc As Document
    Dim nDoc As Document
    Set oDoc = Word.ActiveDocument
    Dim i As Long
    Dim j As Long
    j = oDoc.Sections.Count
    For i = 1 To j
        oDoc.Sections(i).Range.Copy   '复制一节内容'
        Set nDoc = Word.Documents.Add    '新建文档'
        nDoc.Content.PasteAndFormat (wdFormatOriginalFormatting) '按源格式粘贴到新文档
        nDoc.PageSetup = oDoc.Sections(i).PageSetup   '复制原文档页面设置
        Call 删除分节符
        nDoc.SaveAs2 sPath & "\" & oDoc.Name & "_" & i & ".docx"   '保存
        nDoc.Close   '关闭新文档'
    Next
End Sub

Function 删除分节符()

    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
    .Text = "^b"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Function





