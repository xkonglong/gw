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
'  日期: 2023年2月24日


Sub xkonglong_按页拆分文档()

    '本段代码作用: 选择拆分后要保存的文件夹'
    Dim dia As FileDialog
    Set dia = Application.FileDialog(msoFileDialogFolderPicker)
    dia.Title = "请选择拆分后要保存的文件夹"
    If Word.ActiveDocument.Saved Then dia.InitialFileName = Word.ActiveDocument.path
    If dia.Show = 0 Then Exit Sub
    Dim sPath As String
    sPath = dia.SelectedItems(1)
    
    
    
    Dim oDoc As Document
    Dim oRng As Range
    Dim oDocTemp As Document
    Set oDoc = Word.ActiveDocument
    
    Dim iPageNo As Long
    '获取总页数
    With oDoc
    iPageNo = .Range.Information(wdNumberOfPagesInDocument)
        For i = 1 To iPageNo
            '定位到页开始
            Set oRng = .GoTo(wdGoToPage, Which:=wdGoToAbsolute, Count:=i)
            oRng.Select
            '定位整个页面区域
            oRng.SetRange oRng.Start, oRng.Bookmarks("\page").End
            oRng.Copy
            '新建文档粘贴、保存、关闭'
            Set oDocTemp = Word.Documents.Add
            With oDocTemp.Application.Selection
                .Paste
            End With
            oDocTemp.SaveAs2 sPath & "\" & oDoc.Name & "_" & i & ".docx"
            oDocTemp.Close
        Next i
    End With
End Sub
