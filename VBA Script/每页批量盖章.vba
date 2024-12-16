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
'日期: 2024年11月10日

Sub 批量盖章()
Attribute 批量盖章.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏1"

    Dim doc As Document
    Set doc = ActiveDocument
        
    Dim 筛选器 As FileDialog, 文件名, 文件 As Document, pd, 计数 As Integer
    Set 筛选器 = Application.FileDialog(msoFileDialogFilePicker)
    With 筛选器
        .Title = "请指定公章图片"
        .Filters.Add "公章图片", "*.gif; *.png", 1
        .Show
    End With
    If 筛选器.SelectedItems.Count <= 0 Then Exit Sub     '点击取消按钮退出
    
    Dim picPath As String
    picPath = 筛选器.SelectedItems(1)
    
    
    
    ' 遍历文档的所有节
    For Each sec In doc.Sections
        ' 添加一个水印形状
        Set shp = sec.Headers(wdHeaderFooterPrimary).Shapes.AddPicture(picPath)
        
        ' 调整图片大小和位置
        With shp
            .LockAspectRatio = msoTrue
            .Width = CentimetersToPoints(3) ' 宽度设为3厘米
            .Height = CentimetersToPoints(3) ' 高度设为3厘米
            .WrapFormat.Type = wdWrapBehind ' 图片置于文字下方
            .Left = 350  '设置公章的位置, 不要超过页面宽度
            .Top = 500   '设置公章的位置, 不要超过页面高度'
        End With
    Next sec
    
    
    
End Sub
