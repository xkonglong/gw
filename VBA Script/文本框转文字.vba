Sub ConvertTextBoxesToText()
    Dim shp As Shape
    Dim rng As Range
    Dim textToInsert As String
    
   '确保有内容被选中
    If Selection.Type <> wdSelectionShape Then
        MsgBox "请先选择文本框。"
        Exit Sub
    End If
    
   '遍历选中的文本框
    For Each shp In Selection.ShapeRange
        If shp.Type = msoTextBox Then
           '提取文本框中的文本
            textToInsert = shp.TextFrame.TextRange.Text
           '删除文本框
            shp.Delete
           '将文本插入到文档中
            If rng Is Nothing Then
                Set rng = ActiveDocument.Range(Selection.Start, Selection.Start)
            Else
                Set rng = ActiveDocument.Range(rng.End, rng.End)
            End If
            rng.InsertAfter textToInsert
            rng.Collapse wdCollapseEnd
        End If
    Next shp
End Sub