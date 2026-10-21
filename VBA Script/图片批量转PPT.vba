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
'日期:2025年10月31日'

Sub 图片批量转PPT()
	Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim i As Integer
    Dim imagePath As String
    Dim savePath As String
    Dim slideWidth As Single
    Dim slideHeight As Single
    
    ' 创建PowerPoint应用程序对象
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    ' 让PowerPoint可见（可选）
    pptApp.Visible = True
    
    ' 创建新的演示文稿
    Set pptPres = pptApp.Presentations.Add
    
    ' 获取幻灯片尺寸
    With pptPres.PageSetup
        slideWidth = .slideWidth
        slideHeight = .slideHeight
    End With
    
    ' 选择图片文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要转换为PPTX的图片文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff"
        
        If .Show = -1 Then ' 用户选择了文件
            i = 0
            For Each fileItem In .SelectedItems
                imagePath = CStr(fileItem)
                
                ' 添加新幻灯片（使用空白版式）
                Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12) ' ppLayoutBlank
                
                ' 添加图片到幻灯片，填充整个幻灯片
                With pptSlide.Shapes.AddPicture( _
                    FileName:=imagePath, _
                    LinkToFile:=msoFalse, _
                    SaveWithDocument:=msoTrue, _
                    Left:=0, _
                    Top:=0, _
                    Width:=slideWidth, _
                    Height:=slideHeight)
                End With
                
                i = i + 1
            Next fileItem
        Else
            MsgBox "没有选择任何文件。"
            GoTo CleanUp
        End If
    End With
    
CleanUp:
    ' 清理对象
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    Set fd = Nothing
End Sub