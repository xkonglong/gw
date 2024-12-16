'====================================================================
'                            使用说明:  
'制作 VBA 文件时, 请只保留一个Sub 用于执行,  其他过程请用 Function 关键字。
'Sub 过程名推荐用 用户ID_中文名称 的方式命名,避免代码冲突。
'VBA 文件请使用ANSI(简体中文GB2312)编码保存, 微软的VBA解释器不支持UTF-8这类编码,会导致中文乱码。
'不正确的代码可能导致word崩溃、文档丢失或损坏。
'请务必保证vba代码来源安全可靠，插件作者不承担vba文件造成的任何损失！
'分享 VBA 文件时，请尊重作者版权，注明来源以示感谢。
'====================================================================
'原作者: Extendoffice
'修改: 小恐龙  



Sub ExtractBookmarksInADoc()

    Dim xRow As Long
    Dim xTable As Table
    Dim xDoc As Document
    Dim xBookMark As Bookmark
    Dim xBookMarkDoc As Document
    Dim xParagraph As Paragraph
    On Error Resume Next
    Set xDoc = ActiveDocument
    If xDoc.Bookmarks.Count = 0 Then
        MsgBox "文档中没有书签", vbInformation, "VBA脚本:打印书签"
        Exit Sub
    End If
    Set xBookMarkDoc = Documents.Add
    xRow = 1
    Selection.TypeText "书签在 " & "'" & xDoc.Name & "'"
    Set xTable = Selection.Tables.Add(Selection.Range, 1, 3)
    xTable.Borders.Enable = True
    With xTable
        .Cell(xRow, 1).Range.Text = "名称"
        .Cell(xRow, 2).Range.Text = "文字"
        .Cell(xRow, 3).Range.Text = "页码"
        For Each xBookMark In xDoc.Bookmarks
            xTable.Rows.Add
            xRow = xRow + 1
            .Cell(xRow, 1).Range.Text = xBookMark.Name
            .Cell(xRow, 2).Range.Text = xBookMark.Range.Text
            .Cell(xRow, 3).Range.Text = xBookMark.Range.Information(wdActiveEndAdjustedPageNumber)
            xDoc.Hyperlinks.Add Anchor:=.Cell(xRow, 3).Range, Address:=xDoc.Name, _
              SubAddress:=xBookMark.Name, TextToDisplay:=.Cell(xRow, 3).Range.Text
        Next
    End With
    xBookMarkDoc.SaveAs xDoc.Path & "\" & "书签在 " & xDoc.Name
    xBookMarkDoc.PrintOut
    xBookMarkDoc.Close
    Kill xBookMarkDoc.Path
End Sub