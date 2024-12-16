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
'日期:2023年2月2日'

'使用方法:
'运行本vba脚本, 选择多个图片(会自动顺序),然后等待自动插入和保存对话框


Sub 创建图片PDF()

    Dim 筛选器 As FileDialog, 文件名, 文件 As Document, pd, 计数 As Integer
    Set 筛选器 = Application.FileDialog(msoFileDialogFilePicker)
    With 筛选器
        .AllowMultiSelect = True
        .Title = "请选择要批量插入的图片,可多选"
        .Show
    End With
    If 筛选器.SelectedItems.Count <= 0 Then Exit Sub     '点击取消按钮退出
    
    
    Dim doc As Document
    Set doc = Application.Documents.Add()
    
    With doc.PageSetup
        .PaperSize = wdPaperA4 '设为A4纸张
        .LeftMargin = 0   '设置左边距为0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        
    End With
    
    Dim pic As InlineShape
    
    For Each 文件名 In 筛选器.SelectedItems
        
        Set pic = Selection.InlineShapes.AddPicture(文件名)
        pic.Height = doc.PageSetup.PageHeight
        pic.Width = doc.PageSetup.PageWidth
        
    Next
    
  
    
    
    With Application.FileDialog(msoFileDialogSaveAs)
        .InitialFileName = 筛选器.SelectedItems.Item(1)
        .FilterIndex = 7   '7为pdf,个别版本会有不同,请自行调整'
        If (.Show = -1) Then
            .Execute
            doc.Close (False)
        End If
    End With
    
    

End Sub

