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
'日期:2023年2月13日'


Sub 批量打印WORD文档()

    On Error Resume Next

    Dim 筛选器 As FileDialog, 文件名, 文件 As Document, pd, 计数 As Integer
    Set 筛选器 = Application.FileDialog(msoFileDialogFilePicker)
    With 筛选器
        .AllowMultiSelect = True
        .Title = "请选择要批量打印的文档,可多选"
        .Show
    End With
    If 筛选器.SelectedItems.Count <= 0 Then Exit Sub     '点击取消按钮退出
    
    Dim 份数 As Long
    份数 = InputBox("请输入要打印的份数?", "小恐龙VBA脚本", 1)
    
    If 份数 <= 0 Then Exit Sub    '点击取消或份数<=0退出'
    
    For Each 文件名 In 筛选器.SelectedItems
        If Not Right(文件名, Len(文件名) - InStrRev(文件名, ".")) Like "doc*" Then GoTo 结束
        Set 文件 = Documents.Open(文件名)  '打开文件
        文件.PrintOut , , , , , , , 份数 '文件打印
        文件.Close False  '文件关闭
        计数 = 计数 + 1
        Set 文件 = Nothing
        
结束:  Next
    Set 筛选器 = Nothing
    MsgBox "已完成！共打印了" & 计数 & "个文件,各" & 份数 & "份!"
End Sub
